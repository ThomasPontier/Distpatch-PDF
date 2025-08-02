"""
Design tokens and helper utilities for PySide6 UI styling.

Brand primary color: rgb(15, 5, 107)
"""

from PySide6.QtGui import QPalette, QColor
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

# Color palette (tokens)
PRIMARY_RGB = (15, 5, 107)            # brand primary
PRIMARY = QColor(*PRIMARY_RGB)
PRIMARY_600 = QColor(12, 4, 96)
PRIMARY_700 = QColor(10, 3, 86)
PRIMARY_800 = QColor(8, 3, 72)

SURFACE = QColor(250, 250, 250)       # #FAFAFA
SURFACE_ELEVATED = QColor(255, 255, 255)
TEXT_PRIMARY = QColor(34, 34, 34)     # #222
TEXT_MUTED = QColor(68, 68, 68)       # #444
BORDER = QColor(224, 224, 224)        # #E0E0E0
DIVIDER = QColor(221, 221, 221)       # #DDDDDD
SUCCESS = QColor(46, 125, 50)
WARNING = QColor(230, 81, 0)
ERROR = QColor(183, 28, 28)
DISABLED_BG = QColor(158, 158, 158)
DISABLED_TEXT = QColor(224, 224, 224)

# Typography
FONT_FAMILY = 'Segoe UI, Arial, sans-serif'
FONT_SIZE_PT = 12

# Radii
RADIUS_SM = 4
RADIUS_MD = 6
RADIUS_LG = 10

# Spacing scale (px)
SPACE_1 = 4
SPACE_2 = 6
SPACE_3 = 8
SPACE_4 = 10
SPACE_5 = 12
SPACE_6 = 16
SPACE_8 = 20
SPACE_10 = 24

def apply_palette(app: QApplication):
    """
    Apply a light palette aligned with tokens.
    This augments QSS styling with sane widget defaults.
    """
    pal = app.palette()

    pal.setColor(QPalette.Window, SURFACE)
    pal.setColor(QPalette.Base, SURFACE_ELEVATED)
    pal.setColor(QPalette.AlternateBase, SURFACE)
    pal.setColor(QPalette.Text, TEXT_PRIMARY)
    pal.setColor(QPalette.WindowText, TEXT_PRIMARY)
    pal.setColor(QPalette.Button, SURFACE_ELEVATED)
    pal.setColor(QPalette.ButtonText, TEXT_PRIMARY)
    pal.setColor(QPalette.ToolTipBase, SURFACE_ELEVATED)
    pal.setColor(QPalette.ToolTipText, TEXT_PRIMARY)
    pal.setColor(QPalette.Highlight, PRIMARY)
    pal.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    pal.setColor(QPalette.Disabled, QPalette.Text, DISABLED_BG)
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, DISABLED_BG)

    app.setPalette(pal)

def load_qss(app: QApplication, qss: str):
    """
    Set application-wide stylesheet.
    """
    app.setStyleSheet(qss)
