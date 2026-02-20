import numpy as np
import math
import cv2

class Calibrator:
    def __init__(self):
        self.pixel_points = [] # List of (x, y) tuples
        self.graph_points = [] # List of (x, y) tuples (data values)
        
        # Computed scaling factors
        self.x_scale = 1.0
        self.y_scale = 1.0
        self.x_offset = 0.0
        self.y_offset = 0.0
        
        # Log scale flags
        self.is_log_x = False
        self.is_log_y = False
        
        # Cached calibration points for helper access
        self.px_x1 = 0
        self.px_x2 = 1
        self.px_y1 = 0
        self.px_y2 = 1
        self.val_x1 = 0
        self.val_x2 = 1
        self.val_y1 = 0
        self.val_y2 = 1
        
        # Homography Matrix
        self.homography_matrix = None

    def is_calibrated(self):
        return len(self.pixel_points) == 4

    def set_calibration(self, pixel_points, graph_values, is_log_x=False, is_log_y=False):
        """
        pixel_points: List of 4 (x, y) tuples from screen [X1, X2, Y1, Y2]
        graph_values: List of 4 (x, y) tuples from user input
        """
        if len(pixel_points) != 4 or len(graph_values) != 4:
            raise ValueError("Need exactly 4 points for calibration")
        
        self.pixel_points = pixel_points
        self.graph_points = graph_values
        self.is_log_x = is_log_x
        self.is_log_y = is_log_y
        self.homography_matrix = None # Reset if standard calibration is used
        
        # Extract points
        # Point 0: X1 axis start
        # Point 1: X2 axis end
        # Point 2: Y1 axis start
        # Point 3: Y2 axis end
        
        self.px_x1 = pixel_points[0][0]
        self.px_x2 = pixel_points[1][0]
        self.px_y1 = pixel_points[2][1]
        self.px_y2 = pixel_points[3][1]
        
        self.val_x1 = graph_values[0][0]
        self.val_x2 = graph_values[1][0]
        self.val_y1 = graph_values[2][1]
        self.val_y2 = graph_values[3][1]
        
        # Handle Log Scale Input: Convert values to log for linear interpolation
        v_x1 = math.log10(self.val_x1) if self.is_log_x and self.val_x1 > 0 else self.val_x1
        v_x2 = math.log10(self.val_x2) if self.is_log_x and self.val_x2 > 0 else self.val_x2
        v_y1 = math.log10(self.val_y1) if self.is_log_y and self.val_y1 > 0 else self.val_y1
        v_y2 = math.log10(self.val_y2) if self.is_log_y and self.val_y2 > 0 else self.val_y2
        
        # Compute Scale (Slope) and Offset (Intercept)
        # Value = Offset + Pixel * Scale
        # Scale = (V2 - V1) / (P2 - P1)
        
        dx_px = self.px_x2 - self.px_x1
        dy_px = self.px_y2 - self.px_y1
        
        # Avoid division by zero
        if abs(dx_px) < 1e-5: dx_px = 1e-5
        if abs(dy_px) < 1e-5: dy_px = 1e-5
            
        self.x_scale = (v_x2 - v_x1) / dx_px
        self.y_scale = (v_y2 - v_y1) / dy_px
        
        # Offset = V1 - P1 * Scale
        # Offset = V1 - P1 * Scale
        self.x_offset = v_x1 - self.px_x1 * self.x_scale
        self.y_offset = v_y1 - self.px_y1 * self.y_scale

    def set_perspective_calibration(self, px_points, val_points, is_log_x=False, is_log_y=False):
        """
        px_points: List of 4 (x, y) tuples (Corners: TL, TR, BR, BL order doesn't strictly matter as long as matches val_points)
        val_points: List of 4 (x, y) tuples (Data values)
        """
        self.is_log_x = is_log_x
        self.is_log_y = is_log_y
        
        # 1. Process Values (Log Scale)
        processed_dst = []
        for x, y in val_points:
             vx = math.log10(x) if is_log_x and x > 0 else x
             vy = math.log10(y) if is_log_y and y > 0 else y
             processed_dst.append([vx, vy])
             
        src = np.array(px_points, dtype=np.float32).reshape(-1, 1, 2)
        dst = np.array(processed_dst, dtype=np.float32).reshape(-1, 1, 2)
        
        self.homography_matrix, _ = cv2.findHomography(src, dst)

    def map_to_data(self, px, py):
        if self.homography_matrix is not None:
             # Perspective Mapping
             src = np.array([[[px, py]]], dtype=np.float32)
             dst = cv2.perspectiveTransform(src, self.homography_matrix)
             x_val_linear = dst[0][0][0]
             y_val_linear = dst[0][0][1]
        else:
             # Standard Linear Mapping
             x_val_linear = self.x_offset + px * self.x_scale
             y_val_linear = self.y_offset + py * self.y_scale
        
        # 2. Convert back if log scale
        if self.is_log_x:
            try:
                x_data = 10 ** x_val_linear
            except OverflowError:
                x_data = float('inf')
        else:
            x_data = x_val_linear
            
        if self.is_log_y:
            try:
                y_data = 10 ** y_val_linear
            except OverflowError:
                y_data = float('inf')
        else:
            y_data = y_val_linear
        
        return float(x_data), float(y_data)

        self.homography_matrix, _ = cv2.findHomography(src, dst)

    def to_dict(self):
        return {
            'pixel_points': self.pixel_points,
            'graph_points': self.graph_points,
            'is_log_x': self.is_log_x,
            'is_log_y': self.is_log_y,
            'is_perspective': self.homography_matrix is not None
        }

    def from_dict(self, data):
        pixel_points = data.get('pixel_points', [])
        graph_points = data.get('graph_points', [])
        
        # Check if we have valid calibration data
        if not pixel_points or len(pixel_points) != 4 or not graph_points or len(graph_points) != 4:
            # Not fully calibrated or empty project. Just reset/skip.
            self.pixel_points = pixel_points
            self.graph_points = graph_points
            return

        if data.get('is_perspective'):
            self.set_perspective_calibration(
                pixel_points,
                graph_points,
                data.get('is_log_x', False),
                data.get('is_log_y', False)
            )
        else:
             self.set_calibration(
                pixel_points,
                graph_points,
                data.get('is_log_x', False),
                data.get('is_log_y', False)
            )
