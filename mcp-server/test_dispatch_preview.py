import pytest
from unittest.mock import patch, mock_open, AsyncMock
from mcp.types import CallToolResult, TextContent, ImageContent
from server import dispatch_preview

# Global mock setups can be added here or defined in tests.

@pytest.mark.asyncio
@patch("server.ymm4_get", new_callable=AsyncMock)
async def test_dispatch_preview_capture(mock_ymm4_get):
    # Test successful capture
    mock_ymm4_get.return_value = {"success": True, "image": "base64data", "width": 1920, "height": 1080}
    args = {"action": "capture"}
    result = await dispatch_preview(args)

    mock_ymm4_get.assert_called_once_with("/preview/capture")
    assert isinstance(result, CallToolResult)
    assert not result.isError
    assert len(result.content) == 2
    assert isinstance(result.content[0], ImageContent)
    assert result.content[0].data == "base64data"
    assert isinstance(result.content[1], TextContent)

@pytest.mark.asyncio
@patch("server.ymm4_get", new_callable=AsyncMock)
async def test_dispatch_preview_capture_with_element(mock_ymm4_get):
    # Test capture with element param
    mock_ymm4_get.return_value = {"success": True, "image": "base64data", "width": 1920, "height": 1080}
    args = {"action": "capture", "element": "my-element"}
    result = await dispatch_preview(args)

    mock_ymm4_get.assert_called_once_with("/preview/capture?element=my-element")
    assert not result.isError

@pytest.mark.asyncio
@patch("server.ymm4_get", new_callable=AsyncMock)
async def test_dispatch_preview_capture_failure(mock_ymm4_get):
    # Test capture failure
    mock_ymm4_get.return_value = {"success": False, "error": "Capture failed"}
    args = {"action": "capture"}
    result = await dispatch_preview(args)

    assert result.isError
    assert len(result.content) == 1
    assert isinstance(result.content[0], TextContent)
    assert "Capture failed" in result.content[0].text

@pytest.mark.asyncio
@patch("server.ymm4_post", new_callable=AsyncMock)
async def test_dispatch_preview_seek_capture(mock_ymm4_post):
    mock_ymm4_post.return_value = {"success": True, "image": "base64data"}
    args = {"action": "seek_capture", "frame": 100}
    result = await dispatch_preview(args)

    mock_ymm4_post.assert_called_once_with("/preview/seek", {"frame": 100})
    assert not result.isError
    assert isinstance(result.content[0], ImageContent)
    assert isinstance(result.content[1], TextContent)

@pytest.mark.asyncio
@patch("server.ymm4_get", new_callable=AsyncMock)
async def test_dispatch_preview_position(mock_ymm4_get):
    mock_ymm4_get.return_value = {"current_frame": 10, "max_frame": 100}
    args = {"action": "position"}
    result = await dispatch_preview(args)

    mock_ymm4_get.assert_called_once_with("/preview/position")
    assert not result.isError
    assert isinstance(result.content[0], TextContent)
    assert "10" in result.content[0].text

@pytest.mark.asyncio
@patch("server.ymm4_post", new_callable=AsyncMock)
@patch("builtins.open", new_callable=mock_open)
async def test_dispatch_preview_record_success(mock_file, mock_ymm4_post):
    # Test record with audio data
    import base64
    fake_audio = base64.b64encode(b"fake audio data").decode("utf-8")
    mock_ymm4_post.return_value = {
        "success": True,
        "audio": fake_audio,
        "has_audio": True,
        "rms_level": 0.5
    }
    args = {"action": "record", "duration_ms": 2000}
    result = await dispatch_preview(args)

    mock_ymm4_post.assert_called_once_with("/preview/record", {"duration_ms": 2000})
    assert not result.isError
    # Check that file write was attempted
    mock_file.assert_called_once()
    mock_file().write.assert_called_once_with(b"fake audio data")

    assert len(result.content) == 2
    assert isinstance(result.content[0], TextContent)
    assert isinstance(result.content[1], TextContent)
    assert "録音完了" in result.content[1].text
    assert "音声あり ✅" in result.content[1].text

@pytest.mark.asyncio
@patch("server.ymm4_post", new_callable=AsyncMock)
async def test_dispatch_preview_record_no_audio(mock_ymm4_post):
    # Test record when no audio data is returned
    mock_ymm4_post.return_value = {
        "success": True,
        # 'audio' key not present or None
        "message": "Silent recording"
    }
    args = {"action": "record"}
    result = await dispatch_preview(args)

    mock_ymm4_post.assert_called_once_with("/preview/record", {"duration_ms": 3000})
    assert not result.isError
    # Content should just be the format_result text
    assert len(result.content) == 1
    assert "Silent recording" in result.content[0].text

@pytest.mark.asyncio
@patch("server.ymm4_post", new_callable=AsyncMock)
@patch("builtins.open", new_callable=mock_open)
async def test_dispatch_preview_watch_success(mock_file, mock_ymm4_post):
    import base64
    fake_audio = base64.b64encode(b"fake audio").decode("utf-8")
    mock_ymm4_post.return_value = {
        "success": True,
        "start_frame": 0,
        "duration_ms": 5000,
        "audio": {
            "rms_level": 0.1,
            "has_audio": True,
            "data": fake_audio
        },
        "frames": [
            {"time_ms": 0, "image": "frame0data"},
            {"time_ms": 1000, "image": "frame1data"}
        ]
    }
    args = {"action": "watch", "frame": 0, "duration_ms": 5000, "capture_interval_ms": 1000}
    result = await dispatch_preview(args)

    mock_ymm4_post.assert_called_once_with("/preview/watch", {
        "frame": 0,
        "duration_ms": 5000,
        "capture_interval_ms": 1000
    })
    assert not result.isError
    # 1 intro text + 1 save text + (2 * 2 for frames (text + image)) = 6 items
    assert len(result.content) == 6
    assert isinstance(result.content[0], TextContent)
    assert "watch: frame=0 duration=5000ms" in result.content[0].text
    assert isinstance(result.content[1], TextContent)
    assert "音声保存:" in result.content[1].text

    assert isinstance(result.content[2], TextContent)
    assert "t=0ms" in result.content[2].text
    assert isinstance(result.content[3], ImageContent)
    assert result.content[3].data == "frame0data"

@pytest.mark.asyncio
@patch("server.ymm4_post", new_callable=AsyncMock)
async def test_dispatch_preview_watch_failure(mock_ymm4_post):
    mock_ymm4_post.return_value = {
        "success": False,
        "error": "Watch error"
    }
    args = {"action": "watch"}
    result = await dispatch_preview(args)

    assert result.isError
    assert len(result.content) == 1
    assert "Watch error" in result.content[0].text
