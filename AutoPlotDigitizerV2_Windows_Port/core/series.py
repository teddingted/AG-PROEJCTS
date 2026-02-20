from PySide6.QtGui import QColor
import csv

class Series:
    def __init__(self, name: str, color: QColor = None):
        self.name = name
        self.color = color if color else QColor(255, 0, 0)
        self.raw_pixels = [] # List of (x, y) tuples
        self.data_points = [] # List of (x, y) tuples (scaled)
        self.line_type = 'auto' # 'auto', 'manual', 'solid'
        self.gap_fill = 3

    def set_data(self, pixels, data_points):
        self.raw_pixels = pixels
        self.data_points = data_points

    def calculate_instant_gradients(self):
        """
        Calculates instantaneous gradient for each point based on 
        slope between x-0.5 and x+0.5 units.
        Returns a list of float gradients corresponding to data_points.
        """
        if not self.data_points or len(self.data_points) < 2:
            return [0.0] * len(self.data_points)
            
        import numpy as np
        
        # Sort by X just in case
        sorted_points = sorted(self.data_points, key=lambda p: p[0])
        x_vals = np.array([p[0] for p in sorted_points])
        y_vals = np.array([p[1] for p in sorted_points])
        
        grads = []
        for i, (px, py) in enumerate(self.data_points):
            # We want gradient at px. 
            # Define window: [px - 0.5, px + 0.5]
            x_left = px - 0.5
            x_right = px + 0.5
            
            # Interpolate
            # np.interp uses the boundary value for outside range (flat extrapolation)
            y_left = np.interp(x_left, x_vals, y_vals)
            y_right = np.interp(x_right, x_vals, y_vals)
            
            # Run = 1.0 (x_right - x_left)
            rise = y_right - y_left
            grad = rise # / 1.0
            grads.append(grad)
            
        return grads

    def __repr__(self):
        return f"<Series '{self.name}' ({len(self.data_points)} points)>"

    def to_dict(self):
        return {
            'name': self.name,
            'color': self.color.name(),
            'raw_pixels': self.raw_pixels,
            'data_points': self.data_points,
            'line_type': self.line_type,
            'gap_fill': self.gap_fill
        }

    @classmethod
    def from_dict(cls, data):
        series = cls(data['name'], QColor(data['color']))
        series.raw_pixels = data['raw_pixels']
        series.data_points = data['data_points']
        series.line_type = data.get('line_type', 'auto')
        series.gap_fill = data.get('gap_fill', 3)
        return series
