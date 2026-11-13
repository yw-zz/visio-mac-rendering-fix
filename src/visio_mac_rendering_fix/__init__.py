"""Fix Visio / EMF+ figures in Office files (Mac rendering, PDF picture resolution)."""
__version__ = '0.3.0'

from .emf import fix_emf, is_dual_emf
from .fonts import FontIndex
from .package import fix_package

__all__ = ['fix_emf', 'is_dual_emf', 'FontIndex', 'fix_package', '__version__']
