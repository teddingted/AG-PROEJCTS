from PySide6.QtWidgets import QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QGraphicsEllipseItem, QGraphicsPathItem, QGraphicsLineItem, QGraphicsRectItem, QGraphicsTextItem
from PySide6.QtGui import QPixmap, QPainter, QPen, QColor, QImage, QPainterPath, QFont
from PySide6.QtCore import Qt, QPointF, Signal, QRectF

class ImageCanvas(QGraphicsView):
    # Signals to notify MainWindow
    calibration_point_added = Signal(int, float, float) # index, x, y
    point_picked = Signal(float, float) # x, y (pixel coords)
    
    MODE_VIEW = 0
    MODE_CALIBRATE = 1
    MODE_CALIBRATE = 1
    MODE_MASK = 2
    MODE_PICK_POINT = 3
    MODE_MASK_ERASER = 4

    def __init__(self):
        super().__init__()
        self.scene = QGraphicsScene()
        self.setScene(self.scene)
        self.setRenderHint(QPainter.Antialiasing)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.viewport().setMouseTracking(True)
        
        self.image_item = None
        self.current_image = None
        
        # State
        self.mode = self.MODE_VIEW
        self.zoom_factor = 1.15
        
        # Calibration
        self.calibration_points = [] # List of QGraphicsEllipseItem
        
        # Masking
        self.mask_layer = QGraphicsScene() 
        self.drawing_path = None
        self.current_path_item = None
        self.pen_size = 20
        self.mask_items = []
        
        # Extracted Points
        self.extracted_point_items = []
        
        # Crosshairs
        self._init_crosshairs()

    def _init_crosshairs(self):
        self.crosshair_v = QGraphicsLineItem()
        self.crosshair_h = QGraphicsLineItem()
        pen = QPen(QColor(0, 255, 0, 150), 1, Qt.DashLine)
        self.crosshair_v.setPen(pen)
        self.crosshair_h.setPen(pen)
        self.crosshair_v.setZValue(100) # Always on top
        self.crosshair_h.setZValue(100)
        self.crosshair_v.setVisible(False)
        self.crosshair_h.setVisible(False)
        self.scene.addItem(self.crosshair_v)
        self.scene.addItem(self.crosshair_h)

    def set_mode(self, mode):
        self.mode = mode
        if mode == self.MODE_VIEW:
            self.setDragMode(QGraphicsView.ScrollHandDrag)
            self.setCursor(Qt.OpenHandCursor)
        elif mode == self.MODE_CALIBRATE:
            self.setDragMode(QGraphicsView.NoDrag)
            self.setCursor(Qt.CrossCursor)
        elif mode == self.MODE_MASK:
            self.setDragMode(QGraphicsView.NoDrag)
            self.setCursor(Qt.ArrowCursor) 
        elif mode == self.MODE_PICK_POINT:
            self.setDragMode(QGraphicsView.NoDrag)
            self.setCursor(Qt.CrossCursor)
        elif mode == self.MODE_MASK_ERASER:
            self.setDragMode(QGraphicsView.NoDrag)
            self.setCursor(Qt.ForbiddenCursor)
            
        # Hide crosshairs when not calibrating
        if mode != self.MODE_CALIBRATE:
             self.crosshair_v.setVisible(False)
             self.crosshair_h.setVisible(False)

    def load_image(self, file_path):
        self.current_image = QPixmap(file_path)
        self.scene.clear()
        self._init_crosshairs()
        self.calibration_points = []
        self.mask_items = []
        self.extracted_point_items = []
        
        self.image_item = self.scene.addPixmap(self.current_image)
        self.image_item.setZValue(0)
        self.setSceneRect(self.image_item.boundingRect())
        self.fitInView(self.image_item, Qt.KeepAspectRatio)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            # Zoom
            if event.angleDelta().y() > 0:
                self.scale(self.zoom_factor, self.zoom_factor)
            else:
                self.scale(1 / self.zoom_factor, 1 / self.zoom_factor)
            event.accept()
        else:
            super().wheelEvent(event)

    def mousePressEvent(self, event):
        if self.image_item is None:
            return

        sp = self.mapToScene(event.pos())
        
        if self.mode == self.MODE_CALIBRATE:
            if event.button() == Qt.LeftButton:
                if len(self.calibration_points) < 4:
                    self.add_calibration_point(sp)
        
        elif self.mode == self.MODE_MASK or self.mode == self.MODE_MASK_ERASER:
            if event.button() == Qt.LeftButton:
                self.start_drawing(sp)
                
        elif self.mode == self.MODE_PICK_POINT:
            if event.button() == Qt.LeftButton:
                self.point_picked.emit(sp.x(), sp.y())
                
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        sp = self.mapToScene(event.pos())
        
        # Update Crosshairs
        if self.mode == self.MODE_CALIBRATE:
             rect = self.sceneRect()
             self.crosshair_v.setLine(sp.x(), rect.top(), sp.x(), rect.bottom())
             self.crosshair_h.setLine(rect.left(), sp.y(), rect.right(), sp.y())
             self.crosshair_v.setVisible(True)
             self.crosshair_h.setVisible(True)
        else:
             self.crosshair_v.setVisible(False)
             self.crosshair_h.setVisible(False)
        
        if (self.mode == self.MODE_MASK or self.mode == self.MODE_MASK_ERASER) and self.drawing_path:
            self.continue_drawing(sp)

        # Tooltip Logic (Snap to Nearest)
        if self.extracted_point_items:
            # Search radius for "Always Visible" feel (e.g. 50px)
            search_radius = 50
            rect = QRectF(sp.x() - search_radius, sp.y() - search_radius, 
                          search_radius * 2, search_radius * 2)
            
            items_near = self.scene.items(rect)
            
            closest_item = None
            min_dist = float('inf')
            
            for item in items_near:
                if item in self.extracted_point_items:
                    # Calculate distance
                    # Item pos is top-left of ellipse? No, scenePos of item.
                    # QGraphicsEllipseItem pos is usually (0,0) if rect set, or pos if setPos used.
                    # We utilize the rect center or the stored data?
                    # let's use the item's mapping to scene center.
                    # Actually item.sceneBoundingRect().center() is reliable.
                    center = item.sceneBoundingRect().center()
                    dist = (center.x() - sp.x())**2 + (center.y() - sp.y())**2
                    
                    if dist < min_dist:
                        min_dist = dist
                        closest_item = item
            
            if closest_item:
                data = closest_item.data(0)
                origin = closest_item.data(1)
                grad_pres = closest_item.data(2)
                
                if data and origin:
                    # Draw tooltip at Cursor, but show data of closest point
                    self.show_highlight_on_item(closest_item) # Optional: Highlight the specific point?
                    self.draw_tooltip(sp, data[0], data[1], origin[0], origin[1], grad_pres)
            else:
                self.hide_tooltip()
                 
        super().mouseMoveEvent(event)

    def show_highlight_on_item(self, item):
        # Optional: draw a ring around the snapped point
        # Not requested but good for UX. For now, just tooltip.
        pass

    def mouseReleaseEvent(self, event):
        if (self.mode == self.MODE_MASK or self.mode == self.MODE_MASK_ERASER) and self.drawing_path:
            self.finish_drawing()
            
        super().mouseReleaseEvent(event)

    def add_calibration_point(self, pos):
        idx = len(self.calibration_points)
        colors = [Qt.red, Qt.red, Qt.blue, Qt.blue] # X1, X2, Y1, Y2
        
        marker = QGraphicsEllipseItem(pos.x()-5, pos.y()-5, 10, 10)
        marker.setBrush(colors[idx])
        marker.setPen(QPen(Qt.white, 2))
        marker.setZValue(10)
        self.scene.addItem(marker)
        self.calibration_points.append(marker)
        
        self.calibration_point_added.emit(idx, pos.x(), pos.y())

    def restore_calibration_points(self, points):
        """Restores calibration points from list of (x, y) tuples."""
        # Clear existing
        for item in self.calibration_points:
            self.scene.removeItem(item)
        self.calibration_points = []
        
        colors = [Qt.red, Qt.red, Qt.blue, Qt.blue] # X1, X2, Y1, Y2
        for idx, (x, y) in enumerate(points):
            if idx >= 4: break
            marker = QGraphicsEllipseItem(x-5, y-5, 10, 10)
            marker.setBrush(colors[idx])
            marker.setPen(QPen(Qt.white, 2))
            marker.setZValue(10)
            self.scene.addItem(marker)
            self.calibration_points.append(marker)

    def clear_calibration_points(self):
        """Clears all calibration points."""
        for item in self.calibration_points:
            self.scene.removeItem(item)
        self.calibration_points = []

    def start_drawing(self, pos):
        self.drawing_path = QPainterPath(pos)
        
        if self.mode == self.MODE_MASK_ERASER:
            # Eraser: Red Color for Visual Feedback
            pen = QPen(QColor(255, 0, 0, 150), self.pen_size, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
            tag = 'eraser'
        else:
            # Mask: Yellow Color
            pen = QPen(QColor(255, 255, 0, 100), self.pen_size, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
            tag = 'mask'
            
        self.current_path_item = QGraphicsPathItem()
        self.current_path_item.setPen(pen)
        self.current_path_item.setZValue(5)
        self.current_path_item.setData(0, tag)
        self.scene.addItem(self.current_path_item)

    def continue_drawing(self, pos):
        self.drawing_path.lineTo(pos)
        self.current_path_item.setPath(self.drawing_path)

    def finish_drawing(self):
        self.mask_items.append(self.current_path_item)
        self.drawing_path = None
        self.current_path_item = None

    def get_mask_image(self):
        """Renders the mask items to a black and white QImage matching the original image size.
        If no mask is drawn, returns a full-white image (use entire image as ROI)."""
        if self.image_item is None:
            return None
            
        rect = self.image_item.boundingRect()
        width = int(rect.width())
        height = int(rect.height())
        
        image = QImage(width, height, QImage.Format_ARGB32)
        
        # If no mask drawn, use entire image as ROI (full white)
        if not self.mask_items:
            image.fill(Qt.white)
            return image.convertToFormat(QImage.Format_Grayscale8)
        
        # Otherwise start transparent and paint mask on top
        image.fill(Qt.transparent)
        
        painter = QPainter(image)
        painter.setRenderHint(QPainter.Antialiasing)
        
        pen_mask = QPen(Qt.white, self.pen_size, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        pen_eraser = QPen(Qt.black, self.pen_size, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        
        painter.setBrush(Qt.NoBrush)
        
        for item in self.mask_items:
            tag = item.data(0)
            if tag == 'eraser':
                painter.setPen(pen_eraser)
                painter.setCompositionMode(QPainter.CompositionMode_Clear)
            else:
                painter.setPen(pen_mask)
                painter.setCompositionMode(QPainter.CompositionMode_SourceOver)
                
            painter.drawPath(item.path())
            
        painter.end()
        
        # Convert to Grayscale (Transparent -> Black, White -> White)
        return image.convertToFormat(QImage.Format_Grayscale8)

    def clear_mask(self):
        """Clears all drawn mask items from the scene."""
        for item in self.mask_items:
            self.scene.removeItem(item)
        self.mask_items = []
        self.drawing_path = None
        if self.current_path_item:
            self.scene.removeItem(self.current_path_item)
            self.current_path_item = None

    def draw_extracted_points(self, points_px, points_data, gradients=None, color=Qt.green):
        """
        Draws extracted points (pixels) on the scene.
        points_px: list of (px, py)
        points_data: list of (dx, dy)
        gradients: list of float (optional)
        """
        if not points_px or not points_data:
            return
            
        # Enhanced Style: Filled Circle with Border
        pen = QPen(Qt.white, 1) # White border for contrast
        brush = QColor(color)
        brush.setAlpha(255) # Opaque
        radius = 3 # Slightly larger
        
        # Origin for this batch (First point of the series)
        origin_dx, origin_dy = points_data[0]
        
        for i, ((px, py), (dx, dy)) in enumerate(zip(points_px, points_data)):
            item = self.scene.addEllipse(px-radius, py-radius, radius*2, radius*2, pen, brush)
            item.setZValue(8) 
            
            # Store data for tooltip
            # Data 0: Current (dx, dy)
            # Data 1: Origin (dx0, dy0)
            item.setData(0, (dx, dy))
            item.setData(1, (origin_dx, origin_dy))
            
            if gradients and i < len(gradients):
                item.setData(2, gradients[i])
            
            self.extracted_point_items.append(item)

    def clear_extracted_points(self):
        """Clears all extracted point items from the scene."""
        for item in self.extracted_point_items:
            # Check if item is still in scene (might be cleared by scene.clear())
            if item.scene() == self.scene:
                self.scene.removeItem(item)
        self.extracted_point_items = []

    def undo_last_mask(self):
        """Removes the last drawn mask path from the scene."""
        if self.mask_items:
            item = self.mask_items.pop()
            if item.scene() == self.scene:
                self.scene.removeItem(item)

    def draw_tooltip(self, pos, dx, dy, origin_dx, origin_dy, grad_pres=None):
        if not hasattr(self, 'tooltip_box'):
            # Initialize tooltip items
            self.tooltip_box = QGraphicsRectItem()
            self.tooltip_box.setBrush(QColor(255, 255, 255, 220))
            self.tooltip_box.setPen(QPen(Qt.black, 1))
            self.tooltip_box.setZValue(200)
            
            self.tooltip_text = QGraphicsTextItem(self.tooltip_box)
            self.tooltip_text.setHtml("")
            self.tooltip_text.setFont(QFont("Arial", 9))
            
            self.scene.addItem(self.tooltip_box)
        
        # Calculate Stats
        # User requested X/Y absolute values instead of dX/dY
        val_x = dx
        val_y = dy
        
        # GRAD INIT (Total Change / Total Run)
        delta_x = dx - origin_dx
        delta_y = dy - origin_dy
        
        if abs(delta_x) < 1e-9:
            grad_init_str = "Inf"
        else:
            grad_init = delta_y / delta_x
            grad_init_str = f"{grad_init:+.4f}"
            
        # GRAD PRES (Instantaneous)
        grad_pres_str = "N/A"
        if grad_pres is not None:
             grad_pres_str = f"{grad_pres:+.4f}"

        x_str = f"{val_x:.4f}"
        y_str = f"{val_y:.4f}"
        
        # Color formatting
        def color_span(val_str):
            if val_str.startswith('-'):
                return f"<span style='color:red;'>{val_str}</span>"
            return f"<span style='color:black;'>{val_str}</span>"
            
        html = f"""
        <div style='font-family: Arial; font-size: 16px; font-weight: bold;'>
        X: {color_span(x_str)}<br>
        Y: {color_span(y_str)}<br>
        전체변화: {color_span(grad_init_str)}<br>
        단위변화: {color_span(grad_pres_str)}
        </div>
        """
        
        self.tooltip_text.setHtml(html.strip())
        
        # Layout
        text_rect = self.tooltip_text.boundingRect()
        padding = 4
        box_rect = text_rect.adjusted(-padding, -padding, padding, padding)
        self.tooltip_box.setRect(box_rect)
        self.tooltip_text.setPos(0, 0) # Relative to box
        
        # Position near mouse
        self.tooltip_box.setPos(pos.x() + 15, pos.y() + 15)
        self.tooltip_box.setVisible(True)

    def hide_tooltip(self):
        if hasattr(self, 'tooltip_box'):
            self.tooltip_box.setVisible(False)
