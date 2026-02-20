import cv2
import numpy as np

class ImageProcessor:
    def __init__(self):
        pass

    def auto_detect_gap(self, mask_img):
        """
        Tries different gap filling values and returns the best one.
        Best = Minimum number of connected components (ideally 1),
        with valid aspect ratio check if possible.
        """
        best_gap = 1
        min_components = float('inf')
        
        # Test range of gap sizes
        test_gaps = [1, 3, 5, 8, 12, 16]
        
        for gap in test_gaps:
            kernel_size = 2 * gap + 1
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
            closed = cv2.morphologyEx(mask_img, cv2.MORPH_CLOSE, kernel)
            
            # Count connected components
            num_labels, labels = cv2.connectedComponents(closed)
            # num_labels includes background (0), so actual components = num_labels - 1
            components = num_labels - 1
            
            if components < min_components:
                min_components = components
                best_gap = gap
            
            # Optimization: If we hit 1 component, that's likely best (closest to valid line)
            if components == 1:
                break
                
        print(f"DEBUG: Auto-detected gap fill: {best_gap} (components: {min_components})")
        return best_gap

    def find_main_path_shortest(self, points, continuous_mode=False):
        """
        Finds the shortest path...
        continuous_mode: If True, returns empty list if no single path found (strict). 
                         If False (default), falls back to largest components.
        """
        if not points or len(points) < 2:
            return points
            
        # 1. Build Graph (Adjacency List)
        # Map (x,y) -> list of neighbors
        pixel_set = set(points)
        adj = {p: [] for p in points}
        
        offsets = [(-1, -1), (-1, 0), (-1, 1),
                   (0, -1),           (0, 1),
                   (1, -1),  (1, 0),  (1, 1)]
                   
        for p in points:
            x, y = p
            for dx, dy in offsets:
                nx, ny = x + dx, y + dy
                neighbor = (nx, ny)
                if neighbor in pixel_set:
                    adj[p].append(neighbor)
                    
        # 2. Find Start and End nodes (Min X and Max X)
        start_node = min(points, key=lambda p: (p[0], p[1]))
        end_node = max(points, key=lambda p: (p[0], p[1]))
        
        if start_node == end_node:
            return points 
            
        # 3. BFS for Shortest Path
        visited = {start_node}
        parent_map = {start_node: None}
        
        found = False
        import collections
        deque = collections.deque([start_node])
        
        while deque:
            current = deque.popleft()
            
            if current == end_node:
                found = True
                break
                
            for neighbor in adj[current]:
                if neighbor not in visited:
                    visited.add(neighbor)
                    parent_map[neighbor] = current
                    deque.append(neighbor)
        
        if found:
            # Reconstruct
            path = []
            curr = end_node
            while curr is not None:
                path.append(curr)
                curr = parent_map[curr]
            return path[::-1] # Reverse to Start->End
        else:
            # Fallback
            if continuous_mode:
                # In continuous mode, we expect a single connected component.
                # If BFS failed, it means start/end are disconnected.
                # Use largest component only (strict).
                return [] 
                
            # Disconnected component?
            # ... (rest of fallback logic) ...
            # Fallback: Disconnected component?
            # If path search fails (graph is disconnected), return the Largest Connected Component 
            # of the SKELETON graph. This filters out small noise (numbers, dots).
            
            # Find all connected components in 'adj'
            visited_all = set()
            components = []
            
            for p in points:
                if p not in visited_all:
                    # BFS for this component
                    comp = []
                    q = collections.deque([p])
                    visited_all.add(p)
                    while q:
                        curr = q.popleft()
                        comp.append(curr)
                        for neighbor in adj[curr]:
                            if neighbor not in visited_all:
                                visited_all.add(neighbor)
                                q.append(neighbor)
                    components.append(comp)
            
            # Return all significant components (e.g. > 10% of largest)
            if not components:
                return points
                
            largest_comp = max(components, key=len)
            max_len = len(largest_comp)
            thresh = max_len * 0.05
            
            final_points = []
            for comp in components:
                if len(comp) > thresh:
                    final_points.extend(comp)
            
            # Sort by X
            final_points.sort(key=lambda p: (p[0], p[1]))
            return final_points
    

    def process_images(self, original_cv_img, mask_cv_img, hsv_range=None, line_type='auto', gap_fill=None, extraction_mode='segmented'):
        """
        line_type: 'solid', 'manual', or 'auto' (default)
        gap_fill: integer (kernel size parameter) or None
        extraction_mode: 'segmented' (broken lines) or 'continuous' (smooth single line)
        """
        # ... (lines 160-210 same)
        if original_cv_img.shape[:2] != mask_cv_img.shape[:2]:
            mask_cv_img = cv2.resize(mask_cv_img, (original_cv_img.shape[1], original_cv_img.shape[0]))

        # 2. Convert to Mask (Binary)
        if len(mask_cv_img.shape) == 3:
            mask_gray = cv2.cvtColor(mask_cv_img, cv2.COLOR_BGR2GRAY)
        else:
            mask_gray = mask_cv_img
            
        _, roi_mask = cv2.threshold(mask_gray, 10, 255, cv2.THRESH_BINARY)
        print(f"DEBUG: ROI Mask NonZero: {cv2.countNonZero(roi_mask)}")
        
        # 3. Apply Mask to Original
        masked_img = cv2.bitwise_and(original_cv_img, original_cv_img, mask=roi_mask)
        
        # 4. Color Segmentation
        hsv = cv2.cvtColor(masked_img, cv2.COLOR_BGR2HSV)
        
        if hsv_range:
            lower, upper = hsv_range
        else:
            # Default: Dark lines
            lower = np.array([0, 0, 0])
            upper = np.array([180, 255, 100]) 
        
        mask_color = cv2.inRange(hsv, lower, upper)
        print(f"DEBUG: Color Mask NonZero: {cv2.countNonZero(mask_color)}")
        
        # Combine with user ROI
        final_mask = cv2.bitwise_and(mask_color, roi_mask)
        print(f"DEBUG: Final Mask NonZero: {cv2.countNonZero(final_mask)}")
        
        
        # 5. Morphological Operations based on Line Type
        if line_type == 'solid':
            # Solid line: Stronger Closing to bridge grid-line gaps (1-2px)
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (3, 3))
            closed_mask = cv2.morphologyEx(final_mask, cv2.MORPH_CLOSE, kernel, iterations=2)
            
        else: 
            # Dotted/Dashed or Auto: Use Closing to connect gaps
            if line_type == 'auto':
                gap_val = self.auto_detect_gap(final_mask)
            else:
                gap_val = int(gap_fill) if gap_fill is not None else 3
            
            kernel_size = 2 * gap_val + 1
            kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_size, kernel_size))
            closed_mask = cv2.morphologyEx(final_mask, cv2.MORPH_CLOSE, kernel)
        
        # 5.5 FILTERING COMPONENTS BASED ON MODE
        if cv2.countNonZero(closed_mask) > 0:
            num_labels, labels, stats, centroids = cv2.connectedComponentsWithStats(closed_mask, connectivity=8)
            # stats: [x, y, w, h, area]
            
            if num_labels > 1:
                # Find max area
                max_area = 0
                for i in range(1, num_labels):
                    if stats[i, cv2.CC_STAT_AREA] > max_area:
                        max_area = stats[i, cv2.CC_STAT_AREA]
                
                component_mask = np.zeros_like(closed_mask)
                
                if extraction_mode == 'continuous':
                    # CONTINUOUS: Keep ONLY the largest component
                    # This removes all other noise/segments
                    for i in range(1, num_labels):
                        if stats[i, cv2.CC_STAT_AREA] == max_area:
                            component_mask[labels == i] = 255
                            break # Only one
                    closed_mask = component_mask
                    
                else:
                    # SEGMENTED (Default): Keep significant components & Bridge
                    min_area_thresh = max_area * 0.01 
                    
                    for i in range(1, num_labels):
                        if stats[i, cv2.CC_STAT_AREA] >= min_area_thresh:
                            component_mask[labels == i] = 255
                            
                    closed_mask = component_mask
                    
                    # 5.6 Bridging (Only for Segmented/Broken)
                    # Directional Dilation (Horizontal) to jump over vertical grid lines
                    kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 1)) 
                    bridged = cv2.morphologyEx(closed_mask, cv2.MORPH_CLOSE, kernel_h)
                    closed_mask = bridged
        
        # 6. Skeletonization
        if hasattr(cv2, 'ximgproc'):
             skeleton = cv2.ximgproc.thinning(closed_mask)
        else:
            # Simple fallback thinning if ximgproc is missing
            skeleton = closed_mask 
        
        # 7. Extract coordinates
        ys, xs = np.nonzero(skeleton)
        # Ensure native int for JSON serialization
        raw_points = [(int(x), int(y)) for x, y in zip(xs, ys)]
        print(f"DEBUG: Raw Points Found: {len(raw_points)}")
        
        # Sort by X primarily, then Y
        points = sorted(raw_points, key=lambda p: (p[0], p[1]))
        
        # 8. Prune Branches / Order Points
        if extraction_mode == 'continuous':
            # CONTINUOUS: Use Shortest Path logic to find single smooth path
            refined_points = self.find_main_path_shortest(points)
        else:
            # SEGMENTED: Just return sorted points (Shortest path might kill valid segments)
            # Or use find_main_path_shortest if it supports returning multiple components?
            # Our modified find_main_path_shortest DOES return multiple components if disconnected!
            # So we can use it for both, but for Segmented it's critical we keep the fallback logic active.
            refined_points = self.find_main_path_shortest(points)
        
        # Filter points to reduce density (remove clumps)
        filtered_points = []
        if refined_points:
            filtered_points.append(refined_points[0])
            last_p = refined_points[0]
            min_dist_sq = 2.0 * 2.0 
            
            for p in refined_points[1:]:
                dist_sq = (p[0] - last_p[0])**2 + (p[1] - last_p[1])**2
                if dist_sq >= min_dist_sq:
                    filtered_points.append(p)
                    last_p = p
                    
        return filtered_points, skeleton


    def resample_points(self, points, mode='raw', param=None):
        """
        Resamples the list of (x, y) pixels.
        mode: 'raw', 'linear' (fixed count), 'key_points' (Douglas-Peucker)
        param: count for 'linear', epsilon for 'key_points'
        """
        if not points or len(points) < 2:
            return points

        if mode == 'raw':
            return points

        elif mode == 'linear':
            # Fixed Count Interpolation
            target_count = int(param) if param else 50
            if target_count < 2: target_count = 2
            
            # Simple linear interpolation based on index/X
            # Since points are sorted by X, we can interpolate X and Y separately
            # over the cumulative distance or just index.
            # Using cumulative distance is better for curves.
            
            x = np.array([p[0] for p in points])
            y = np.array([p[1] for p in points])
            
            # Cumulative distance
            dist = np.cumsum(np.sqrt(np.ediff1d(x, to_begin=0)**2 + np.ediff1d(y, to_begin=0)**2))
            dist = dist / dist[-1] # Normalize 0..1
            
            target_dist = np.linspace(0, 1, target_count)
            
            new_x = np.interp(target_dist, dist, x)
            new_y = np.interp(target_dist, dist, y)
            
            return list(zip(new_x, new_y))

        elif mode == 'key_points':
            # Douglas-Peucker Simplification
            epsilon = float(param) if param else 2.0
            
            # cv2.approxPolyDP expects a contour (N, 1, 2) array
            contour = np.array(points, dtype=np.float32).reshape((-1, 1, 2))
            approx = cv2.approxPolyDP(contour, epsilon, False) # False = open curve
            
            return [(p[0][0], p[0][1]) for p in approx]

        return points
