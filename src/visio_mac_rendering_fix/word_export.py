"""Export a .docx to PDF with Microsoft Word for Mac, without touching the user's own documents.

The document is copied into a fresh folder inside Word's sandbox container (so Word can read
it without a permission prompt), opened under a unique name, exported with tracked-change
markup hidden, and closed again. Only that copy is ever addressed: every step checks that
Word's active document is the copy before acting, so open documents of the user are never
saved, closed or modified. The copy and its folder are deleted afterwards.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import uuid

WORD_CONTAINER = os.path.expanduser('~/Library/Containers/com.microsoft.Word/Data')


class ExportError(RuntimeError):
    pass


def _osascript(script, timeout):
    try:
        r = subprocess.run(['osascript', '-'], input=script, capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise ExportError('Word did not respond within %d s' % timeout)
    if r.returncode != 0:
        msg = r.stderr.strip()
        if '-1743' in msg or 'Not authorized' in msg:
            msg += ('\nAllow your terminal to control Microsoft Word in System Settings > '
                    'Privacy & Security > Automation, then run again.')
        raise ExportError(msg or 'osascript failed')
    return r.stdout.strip()


def _as_str(s):
    return '"' + s.replace('\\', '\\\\').replace('"', '\\"') + '"'


def word_available():
    return sys.platform == 'darwin' and os.path.isdir('/Applications/Microsoft Word.app')


def export_docx_to_pdf(docx, pdf, markup=False, settle=5, timeout=300):
    """Write `pdf` from `docx` using Word for Mac. Returns Word's page count for the document.

    markup=False exports the final text (tracked changes applied, comments hidden), which is
    what a submission PDF needs; markup=True keeps Word's current markup view.
    """
    if not word_available():
        raise ExportError('automatic export needs Microsoft Word for Mac; '
                          'export the PDF from Word yourself and pass it as well')
    os.makedirs(WORD_CONTAINER, exist_ok=True)
    work = tempfile.mkdtemp(prefix='visio-pdf-hires-', dir=WORD_CONTAINER)
    name = 'vph_%s.docx' % uuid.uuid4().hex[:10]           # unique: can't collide with a user document
    src = os.path.join(work, name)
    out = os.path.join(work, name[:-5] + '.pdf')
    shutil.copyfile(docx, src)
    hide = '' if markup else '''
    set v to view of active window
    set show revisions and comments of v to false
    set revisions view of v to revisions view final'''
    script = '''
tell application "Microsoft Word"
  open file name (POSIX file %(src)s as text) without add to recent files
  set ok to false
  repeat 120 times
    try
      if name of active document is %(name)s then
        set ok to true
        exit repeat
      end if
    end try
    delay 0.5
  end repeat
  if not ok then error "Word did not open the temporary copy"
  delay %(settle)d%(hide)s
  repaginate active document
  delay 1
  if name of active document is not %(name)s then error "active document changed during export"
  set pages to compute statistics active document statistic statistic pages
  save as active document file name (POSIX file %(out)s as text) file format format PDF
  if name of active document is %(name)s then close active document saving no
  return pages
end tell
''' % dict(src=_as_str(src), out=_as_str(out), name=_as_str(name), settle=settle, hide=hide)
    try:
        pages = _osascript(script, timeout)
        if not os.path.exists(out):
            raise ExportError('Word did not write the PDF')
        shutil.move(out, pdf)
        try:
            return int(pages)
        except ValueError:
            return None
    finally:
        shutil.rmtree(work, ignore_errors=True)
