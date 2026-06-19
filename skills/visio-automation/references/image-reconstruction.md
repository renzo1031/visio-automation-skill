# Image Reconstruction

Use this file when recreating an image, screenshot, whiteboard, PDF crop, or sketch as editable Visio.

## Reconstruction Workflow

1. Treat the image as a reference, not as the final artifact. The output should be native Visio shapes, text, and glued connectors, with the original image used only for analysis or temporary alignment.
2. Create a coordinate map from image pixels to Visio page coordinates. Choose page width/height or scale first, then convert every detected node center, size, connector endpoint, and waypoint through the same transform.
3. Build an intermediate reconstruction model before drawing:
   - Nodes: `id`, semantic type, bounding box, center, text, fill/line colors, and likely Visio master.
   - Connectors: `from`, `to`, begin/end attachment points, route type, waypoints/control points, arrowhead direction, line color/weight, and label text if any.
   - Free text or annotations: text, bounding box, style, and nearby anchor if it belongs to a connector.
4. Identify shapes by meaning first, appearance second. For example, map databases to `Database`/data-store masters and decision diamonds to `Decision`; use basic rectangles/ellipses only for purely visual boxes or unknown elements.
5. Recreate connector geometry from the reference image:
   - Right-angle connectors: preserve horizontal/vertical segments and visible 90-degree bends; use `Connect-VisioShapesOrthogonal` and keep orthogonal routing.
   - Straight connectors: use `Connect-VisioShapesStraight` and verify `ShapeRouteStyle = 2`, `ConLineRouteExt = 1`.
   - Curved connectors: use `Connect-VisioShapesCurved` and verify `ConLineRouteExt = 2`. If the curve has important control geometry that Visio routing cannot match closely, draw a native curve/freeform only as a last fallback and still glue or clearly align its endpoints.
6. Glue connector endpoints with `GlueToPos` using the relative side/edge position seen in the image, rather than simply drawing lines between absolute points. This keeps arrows moving with shapes after manual edits.
7. When high positional fidelity matters, place the source image as a temporary locked/low-opacity reference layer, draw native objects over it, then remove or hide the reference before final delivery unless the user asks to keep it.
8. Export a preview PNG and compare it with the source image. Adjust obvious coordinate, size, route, and text placement drift before reporting completion.

## Practical Guidance

- Treat OCR/computer vision as an upstream step. This skill expects structured JSON or analyzed coordinates rather than doing image understanding itself.
- If the user specifically wants the source image preserved, keep the temporary reference layer out of the final output unless that would break the requested result.
