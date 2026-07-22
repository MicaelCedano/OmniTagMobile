# Fix Preview Pixelation

The user reports the preview looks pixelated. This is likely due to aliasing during the downscaling of the high-res generated image to the preview widget size, or missing render hints in the painter.

## Todo List
- [x] Improve `ImagePreviewWidget` rendering <!-- id: 16 -->
    - [x] Add `QPainter.SmoothPixmapTransform` render hint <!-- id: 17 -->
- [x] Increase source image resolution <!-- id: 18 -->
    - [x] Bump DPI from 300 to 600 in `generate_pil_image` for sharper text <!-- id: 19 -->
