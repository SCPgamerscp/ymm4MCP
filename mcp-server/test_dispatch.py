import pytest
from unittest.mock import patch
import sys
import os

# Ensure we can import from server.py
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import server

# We will patch ymm4_get and ymm4_post
@pytest.fixture
def mock_ymm4_get():
    with patch('server.ymm4_get') as mock_get:
        mock_get.return_value = {"success": True}
        yield mock_get

@pytest.fixture
def mock_ymm4_post():
    with patch('server.ymm4_post') as mock_post:
        mock_post.return_value = {"success": True}
        yield mock_post

@pytest.fixture
def mock_add_script():
    with patch('server.add_script') as mock_add_script:
        mock_add_script.return_value = {"success": True, "added": 1}
        yield mock_add_script

@pytest.mark.asyncio
async def test_dispatch_get_info(mock_ymm4_get):
    test_cases = [
        ("status", "/status"),
        ("project", "/project"),
        ("items", "/items"),
        ("effects_list", "/effects/list"),
    ]
    for sub_action, expected_path in test_cases:
        mock_ymm4_get.reset_mock()
        result = await server.dispatch({"action": "get_info", "sub_action": sub_action})
        mock_ymm4_get.assert_called_once_with(expected_path)
        assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_get_info_invalid():
    with pytest.raises(ValueError, match="Unknown sub_action for get_info: invalid"):
        await server.dispatch({"action": "get_info", "sub_action": "invalid"})

@pytest.mark.asyncio
async def test_dispatch_control(mock_ymm4_post):
    test_cases = [
        ("play", "/playback/play"),
        ("stop", "/playback/stop"),
        ("save", "/project/save"),
    ]
    for sub_action, expected_path in test_cases:
        mock_ymm4_post.reset_mock()
        result = await server.dispatch({"action": "control", "sub_action": sub_action})
        mock_ymm4_post.assert_called_once_with(expected_path)
        assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_control_invalid():
    with pytest.raises(ValueError, match="Unknown sub_action for control: invalid"):
        await server.dispatch({"action": "control", "sub_action": "invalid"})

@pytest.mark.asyncio
async def test_dispatch_add_item(mock_ymm4_post):
    test_cases = [
        ("text", "/items/text"),
        ("voice", "/items/voice"),
        ("tachie", "/items/tachie"),
        ("face", "/items/face"),
    ]

    args = {
        "text": "Hello",
        "character": "Reimu",
        "frame": 10,
        "layer": 1,
        "length": 30
    }

    for sub_action, expected_path in test_cases:
        mock_ymm4_post.reset_mock()

        # Test with all payload args
        call_args = {"action": "add_item", "sub_action": sub_action}
        call_args.update(args)

        result = await server.dispatch(call_args)
        mock_ymm4_post.assert_called_once_with(expected_path, args)
        assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_add_item_empty_payload(mock_ymm4_post):
    result = await server.dispatch({"action": "add_item", "sub_action": "text"})
    mock_ymm4_post.assert_called_once_with("/items/text", {})
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_add_item_invalid():
    with pytest.raises(ValueError, match="Unknown sub_action for add_item: invalid"):
        await server.dispatch({"action": "add_item", "sub_action": "invalid"})

@pytest.mark.asyncio
async def test_dispatch_edit_item_face_param(mock_ymm4_post):
    args = {
        "action": "edit_item",
        "sub_action": "face_param",
        "params": {"eye": 1, "mouth": 2},
        "frame": 5,
        "layer": 2
    }
    result = await server.dispatch(args)
    mock_ymm4_post.assert_called_once_with("/items/face/param", {"eye": 1, "mouth": 2, "frame": 5, "layer": 2})
    assert result == {"success": True}

    # Test without frame/layer
    mock_ymm4_post.reset_mock()
    args2 = {"action": "edit_item", "sub_action": "face_param", "params": {"eye": 1}}
    result = await server.dispatch(args2)
    mock_ymm4_post.assert_called_once_with("/items/face/param", {"eye": 1})
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_edit_item_property(mock_ymm4_post):
    args = {
        "action": "edit_item",
        "sub_action": "property",
        "frame": 10,
        "layer": 1,
        "prop": "X",
        "value": 100.5
    }
    result = await server.dispatch(args)
    mock_ymm4_post.assert_called_once_with("/items/prop", {
        "frame": 10,
        "layer": 1,
        "prop": "X",
        "value": "100.5" # Note: str(value) conversion
    })
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_edit_item_effect(mock_ymm4_post):
    args = {
        "action": "edit_item",
        "sub_action": "effect",
        "frame": 5,
        "layer": 2,
        "effect": "Blur"
    }
    result = await server.dispatch(args)
    mock_ymm4_post.assert_called_once_with("/items/effect", {
        "frame": 5,
        "layer": 2,
        "effect": "Blur"
    })
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_edit_item_delete(mock_ymm4_post):
    # Test with frame and layer
    args = {
        "action": "edit_item",
        "sub_action": "delete",
        "frame": 10,
        "layer": 1
    }
    result = await server.dispatch(args)
    mock_ymm4_post.assert_called_once_with("/items/delete", {"frame": 10, "layer": 1})
    assert result == {"success": True}

    # Test with -1 (should not be included)
    mock_ymm4_post.reset_mock()
    args2 = {
        "action": "edit_item",
        "sub_action": "delete",
        "frame": -1,
        "layer": -1,
        "layers": [1, 2]
    }
    result = await server.dispatch(args2)
    mock_ymm4_post.assert_called_once_with("/items/delete", {"layers": [1, 2]})
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_edit_item_duration(mock_ymm4_post):
    args = {
        "action": "edit_item",
        "sub_action": "duration",
        "frames": 100
    }
    result = await server.dispatch(args)
    mock_ymm4_post.assert_called_once_with("/timeline/duration", {"frames": 100})
    assert result == {"success": True}

    # Missing frames defaults to 0
    mock_ymm4_post.reset_mock()
    result = await server.dispatch({"action": "edit_item", "sub_action": "duration"})
    mock_ymm4_post.assert_called_once_with("/timeline/duration", {"frames": 0})
    assert result == {"success": True}

@pytest.mark.asyncio
async def test_dispatch_edit_item_invalid():
    with pytest.raises(ValueError, match="Unknown sub_action for edit_item: invalid"):
        await server.dispatch({"action": "edit_item", "sub_action": "invalid"})

@pytest.mark.asyncio
async def test_dispatch_add_script(mock_add_script):
    args = {
        "action": "add_script",
        "lines": [{"character": "Reimu", "text": "Hi"}]
    }
    result = await server.dispatch(args)
    mock_add_script.assert_called_once_with(args)
    assert result == {"success": True, "added": 1}

@pytest.mark.asyncio
async def test_dispatch_invalid_action():
    with pytest.raises(ValueError, match="Unknown action: invalid"):
        await server.dispatch({"action": "invalid"})

@pytest.mark.asyncio
async def test_dispatch_missing_action():
    with pytest.raises(ValueError, match="Unknown action: None"):
        await server.dispatch({})
