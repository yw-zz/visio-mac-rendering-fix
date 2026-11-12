"""Fix Visio / EMF+ figures that render wrong in Office for Mac."""
__version__ = '0.1.0'

from .emf import fix_emf, is_dual_emf
from .fonts import FontIndex
from .package import fix_package

__all__ = ['fix_emf', 'is_dual_emf', 'FontIndex', 'fix_package', '__version__']
