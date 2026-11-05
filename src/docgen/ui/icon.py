"""The application icon, drawn in code (no binary files in the repository).

A blue tile with a white document (the Word file) and a green check mark (filled from Excel).
`ico_bytes()` builds the multi-size Windows .ico used by the executables and the installer
(packaging/make_icon.py); `app_icon()` is the window icon.
"""

from __future__ import annotations

import struct

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QIcon, QImage, QPainter, QPen, QPixmap, QPolygonF

SIZES = (16, 24, 32, 48, 64, 128, 256)
BLUE, WHITE, GREEN, LINE = QColor("#2563eb"), QColor("#ffffff"), QColor("#16a34a"), QColor("#93c5fd")


def paint(size: int) -> QImage:
    image = QImage(size, size, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    s = float(size)
    p = QPainter(image)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(BLUE)
    p.drawRoundedRect(QRectF(0.03 * s, 0.03 * s, 0.94 * s, 0.94 * s), 0.2 * s, 0.2 * s)
    # the document, with a folded corner
    left, top, right, bottom, fold = 0.26 * s, 0.16 * s, 0.74 * s, 0.84 * s, 0.16 * s
    p.setBrush(WHITE)
    p.drawPolygon(
        QPolygonF(
            [
                QPointF(left, top),
                QPointF(right - fold, top),
                QPointF(right, top + fold),
                QPointF(right, bottom),
                QPointF(left, bottom),
            ]
        )
    )
    p.setBrush(LINE)
    p.drawPolygon(
        QPolygonF([QPointF(right - fold, top), QPointF(right, top + fold), QPointF(right - fold, top + fold)])
    )
    if size >= 24:  # text lines are only visible from 24 px
        pen = QPen(LINE, max(1.0, 0.045 * s))
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        p.setPen(pen)
        for i, width in enumerate((0.36, 0.30, 0.36, 0.22)):
            y = top + 0.26 * s + i * 0.11 * s
            p.drawLine(QPointF(left + 0.07 * s, y), QPointF(left + 0.07 * s + width * s, y))
    # the data: a green badge in the lower right corner
    p.setPen(Qt.PenStyle.NoPen)
    p.setBrush(GREEN)
    badge = 0.38 * s
    p.drawRoundedRect(QRectF(s - badge - 0.06 * s, s - badge - 0.06 * s, badge, badge), 0.08 * s, 0.08 * s)
    if size >= 24:  # a check mark: the document is filled
        pen = QPen(WHITE, max(1.6, 0.06 * s))
        pen.setCapStyle(Qt.PenCapStyle.RoundCap)
        pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        x0, y0 = s - badge - 0.06 * s, s - badge - 0.06 * s
        p.drawPolyline(
            QPolygonF(
                [
                    QPointF(x0 + 0.24 * badge, y0 + 0.52 * badge),
                    QPointF(x0 + 0.43 * badge, y0 + 0.71 * badge),
                    QPointF(x0 + 0.77 * badge, y0 + 0.31 * badge),
                ]
            )
        )
    p.end()
    return image


def png_bytes(size: int) -> bytes:
    data = QByteArray()
    buffer = QBuffer(data)
    buffer.open(QIODevice.OpenModeFlag.WriteOnly)
    paint(size).save(buffer, "PNG")
    buffer.close()
    return bytes(data.data())


def ico_bytes(sizes: tuple[int, ...] = SIZES) -> bytes:
    """A Windows .ico with PNG-compressed images (Windows Vista and later)."""
    images = [png_bytes(size) for size in sizes]
    header = struct.pack("<HHH", 0, 1, len(images))
    offset = len(header) + 16 * len(images)
    entries, payload = b"", b""
    for size, png in zip(sizes, images, strict=True):
        dim = 0 if size >= 256 else size
        entries += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(png), offset + len(payload))
        payload += png
    return header + entries + payload


def app_icon() -> QIcon:
    icon = QIcon()
    for size in SIZES:
        icon.addPixmap(QPixmap.fromImage(paint(size)))
    return icon
