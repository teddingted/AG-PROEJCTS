from .calibration import Calibrator
from .series import Series
import csv
import json

class Project:
    def __init__(self):
        self.image_path = None
        self.calibration = Calibrator()
        self.series_list = [] # List[Series]
        self._observers = []

    def set_image(self, path):
        self.image_path = path
        self.clear_data()
        self.notify_observers()

    def clear_data(self):
        self.series_list = []
        # We might want to keep calibration or reset it?
        # Typically new image = new calibration needed, but sometimes same setup.
        # For now, let's keep calibration object but it will be invalid until set.
        self.calibration = Calibrator()
        self.notify_observers()

    def add_series(self, series: Series):
        self.series_list.append(series)
        self.notify_observers()

    def remove_series(self, index: int):
        if 0 <= index < len(self.series_list):
            self.series_list.pop(index)
            self.notify_observers()

    def update_calibration(self, pixel_points, graph_values, is_log_x=False, is_log_y=False):
        self.calibration.set_calibration(pixel_points, graph_values, is_log_x, is_log_y)
        self.notify_observers()

    def update_calibration_perspective(self, pixel_points, graph_values, is_log_x=False, is_log_y=False):
        self.calibration.set_perspective_calibration(pixel_points, graph_values, is_log_x, is_log_y)
        self.notify_observers()


    def get_csv_data(self):
        """Returns header and rows for CSV export"""
        if not self.series_list:
            return [], []

        max_len = max(len(s.data_points) for s in self.series_list)
        header = []
        for s in self.series_list:
            header.append(f"{s.name}_X")
            header.append(f"{s.name}_Y")

        rows = []
        for i in range(max_len):
            row = []
            for s in self.series_list:
                if i < len(s.data_points):
                    row.append(s.data_points[i][0])
                    row.append(s.data_points[i][1])
                else:
                    row.append("")
                    row.append("")
            rows.append(row)
            
        return header, rows

    def save_project(self, filepath):
        data = {
            'image_path': self.image_path,
            'calibration': self.calibration.to_dict(),
            'series_list': [s.to_dict() for s in self.series_list]
        }
        
        # Robust conversion for Numpy types
        serializable_data = self._convert_to_serializable(data)
        
        with open(filepath, 'w') as f:
            json.dump(serializable_data, f, indent=4)

    def _convert_to_serializable(self, obj):
        import numpy as np
        if isinstance(obj, dict):
            return {k: self._convert_to_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._convert_to_serializable(v) for v in obj]
        elif isinstance(obj, tuple):
            return tuple(self._convert_to_serializable(v) for v in obj)
        elif isinstance(obj, (np.int8,
                            np.int16, np.int32, np.int64, np.uint8,
                            np.uint16, np.uint32, np.uint64)):
            return int(obj)
        elif isinstance(obj, (np.float16, np.float32, np.float64)):
            return float(obj)
        elif isinstance(obj, (np.ndarray,)):
            return self._convert_to_serializable(obj.tolist())
        return obj

    def load_project(self, filepath):
        with open(filepath, 'r') as f:
            data = json.load(f)
            
        self.image_path = data.get('image_path')
        self.calibration.from_dict(data.get('calibration', {}))
        
        self.series_list = []
        for s_data in data.get('series_list', []):
            self.series_list.append(Series.from_dict(s_data))
            
        self.notify_observers()

    def add_observer(self, callback):
        self._observers.append(callback)

    def notify_observers(self):
        for callback in self._observers:
            callback()
