import base64
import json
import struct

import pytest
from PySide6.QtCore import QSize
from PySide6.QtWidgets import QDialog, QDialogButtonBox

from folio.structured_media import (
    EquationDialog, MediaError, ModelScene, load_model, parse_equation,
    render_equation, render_model, render_svg,
)


def _painted_pixels(image):
    return sum(image.pixelColor(x, y).alpha() > 0 for y in range(image.height()) for x in range(image.width()))


def _triangle_model():
    data = struct.pack("<9f3H", 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 1, 2)
    model = {
        "asset": {"version": "2.0"},
        "buffers": [{"byteLength": len(data), "uri": "data:application/octet-stream;base64," + base64.b64encode(data).decode()}],
        "bufferViews": [{"buffer": 0, "byteLength": 36}, {"buffer": 0, "byteOffset": 36, "byteLength": 6}],
        "accessors": [{"bufferView": 0, "componentType": 5126, "count": 3, "type": "VEC3"},
                      {"bufferView": 1, "componentType": 5123, "count": 3, "type": "SCALAR"}],
        "meshes": [{"primitives": [{"attributes": {"POSITION": 0}, "indices": 1}]}],
        "nodes": [{"mesh": 0}], "scenes": [{"nodes": [0]}], "scene": 0,
    }
    return model, data


def test_equations_parse_to_nested_structures_and_render(qapp):
    source = r"\frac{x_{1}^{2}+\sqrt[3]{y}}{\sum_{n=1}^{5}n} = \int_{0}^{1}x\,dx"
    ast = parse_equation(source)
    fraction = ast.children[0]
    assert fraction.kind == "fraction"
    assert fraction.children[0].children[0].kind == "scripts"
    assert fraction.children[0].children[2].kind == "root"
    image = render_equation(source)
    assert image.width() > 100 and image.height() > 70
    assert _painted_pixels(image) > 500
    assert image.pixelColor(0, 0).alpha() == 0


def test_mathml_and_tex_share_fraction_tree_and_layout(qapp):
    mathml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>x</mi><msqrt><mi>y</mi></msqrt></mfrac></math>'
    ast = parse_equation(mathml, "mathml")
    assert ast.children[0].kind == "fraction"
    assert ast.children[0].children[1].kind == "root"
    image = render_equation(mathml, "mathml")
    tex_image = render_equation(r"\frac{x}{\sqrt{y}}")
    assert image == tex_image


@pytest.mark.parametrize("source,syntax", [
    (r"\frac{x}", "tex"), ("x^{2", "tex"), (r"\unknown{x}", "tex"),
    ("x^^2", "tex"), ("x_1_2", "tex"), ("", "tex"),
    ("{" * 40 + "x" + "}" * 40, "tex"),
    ("<math><mfrac><mi>x</mi></mfrac></math>", "mathml"),
    ("<math><merror>broken</merror></math>", "mathml"),
    ('<!DOCTYPE math [<!ENTITY x "foo">]><math><mi>&x;</mi></math>', "mathml"),
    ('<math><mi href="https://example.com">x</mi></math>', "mathml"),
])
def test_invalid_equations_fail_explicitly(source, syntax):
    with pytest.raises(MediaError):
        parse_equation(source, syntax)


def test_equation_dialog_preserves_source_and_blocks_invalid_accept(qapp):
    dialog = EquationDialog(source=r"\frac{a}{b}")
    assert dialog.source() == r"\frac{a}{b}"
    assert dialog.rendered_image().width() > 10
    dialog.source_edit.setPlainText(r"\frac{a}")
    assert dialog.error_label.text()
    assert not dialog.buttons.button(QDialogButtonBox.StandardButton.Ok).isEnabled()
    dialog.accept()
    assert dialog.result() != QDialog.DialogCode.Accepted
    with pytest.raises(MediaError):
        dialog.rendered_image()
    dialog.source_edit.setPlainText("x^2")
    dialog.accept()
    assert dialog.result() == QDialog.DialogCode.Accepted


def test_svg_renders_internal_gradient_and_preserves_aspect(qapp):
    svg = b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 50"><defs><linearGradient id="g"><stop stop-color="red"/><stop offset="1" stop-color="blue"/></linearGradient></defs><rect width="100" height="50" fill="url(\'#g\')"/></svg>'
    image = render_svg(svg, QSize(200, 200))
    assert image.size() == QSize(200, 100)
    assert image.pixelColor(10, 10).red() > image.pixelColor(190, 10).red()


@pytest.mark.parametrize("body", [
    '<image href="file:///etc/passwd"/>', '<image href="https://example.com/x.png"/>',
    '<style>@import "x.css";</style>', '<rect fill="url(https://example.com/x)"/>',
    '<script>alert(1)</script>', '<animate attributeName="x"/>', '<rect onclick="go()"/>',
])
def test_svg_rejects_external_or_active_content(body, qapp):
    with pytest.raises(MediaError):
        render_svg(f'<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">{body}</svg>'.encode())


def test_gltf_embedded_geometry_transforms_and_poster(tmp_path, qapp):
    model, _ = _triangle_model()
    model["nodes"] = [{"translation": [5, 2, 0], "children": [1]}, {"mesh": 0, "scale": [2, 3, 1]}]
    path = tmp_path / "triangle.gltf"
    path.write_text(json.dumps(model))
    scene = load_model(path)
    assert scene.vertices == ((5, 2, 0), (7, 2, 0), (5, 5, 0))
    assert scene.triangles == ((0, 1, 2),)
    assert ModelScene.from_dict(json.loads(json.dumps(scene.to_dict()))) == scene
    poster = render_model(scene, QSize(200, 150))
    assert poster.size() == QSize(200, 150)
    assert _painted_pixels(poster) > 1000
    assert poster.pixelColor(0, 0).alpha() == 0


def test_glb_binary_buffer_loads(tmp_path):
    model, binary = _triangle_model()
    del model["buffers"][0]["uri"]
    source = json.dumps(model).encode()
    source += b" " * (-len(source) % 4)
    binary += b"\0" * (-len(binary) % 4)
    payload = struct.pack("<II", len(source), 0x4E4F534A) + source + struct.pack("<II", len(binary), 0x004E4942) + binary
    path = tmp_path / "triangle.glb"
    path.write_bytes(struct.pack("<4sII", b"glTF", 2, len(payload) + 12) + payload)
    assert len(load_model(path).triangles) == 1


def test_gltf_local_buffers_resolve_only_inside_model_directory(tmp_path):
    model, binary = _triangle_model()
    folder = tmp_path / "model"
    folder.mkdir()
    path = folder / "triangle.gltf"
    (folder / "buffer.bin").write_bytes(binary)
    model["buffers"][0]["uri"] = "buffer.bin"
    path.write_text(json.dumps(model))
    assert len(load_model(path).vertices) == 3
    for uri in ("../buffer.bin", "https://example.com/data.bin", "file:///data.bin", "%2e%2e/buffer.bin", "C:%5Cbuffer.bin"):
        model["buffers"][0]["uri"] = uri
        path.write_text(json.dumps(model))
        with pytest.raises(MediaError):
            load_model(path)


@pytest.mark.parametrize("change", [
    lambda model: model["accessors"][0].update(count=999999),
    lambda model: model["accessors"][0].update(byteOffset=1024),
    lambda model: model["accessors"][0].update(bufferView=-1),
    lambda model: model["nodes"][0].update(children=[0]),
    lambda model: model["nodes"][0].update(scale=[float("nan"), 1, 1]),
    lambda model: model["meshes"][0]["primitives"][0].update(mode=1),
    lambda model: model.update(extensionsRequired=["KHR_draco_mesh_compression"]),
    lambda model: model.update(animations=[{}]),
])
def test_invalid_or_unsupported_models_fail_explicitly(change, tmp_path):
    model, _ = _triangle_model()
    change(model)
    path = tmp_path / "invalid.gltf"
    path.write_text(json.dumps(model))
    with pytest.raises(MediaError):
        load_model(path)
