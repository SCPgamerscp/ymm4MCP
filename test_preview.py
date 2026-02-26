import httpx
import asyncio
import base64
import os
import json

BASE_URL = "http://localhost:8765/api"

async def test_capture():
    print("Testing capture...")
    async with httpx.AsyncClient() as client:
        resp = await client.get(f"{BASE_URL}/preview/capture")
        data = resp.json()
        if data["success"]:
            print(f"Capture success: {data['width']}x{data['height']}")
            img_data = base64.b64decode(data["image"])
            with open("capture_result.png", "wb") as f:
                f.write(img_data)
            print("Saved capture_result.png")
        else:
            print(f"Capture failed: {data.get('error')}")

async def test_seek_capture():
    print("Testing seek and capture at frame 100...")
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{BASE_URL}/preview/seek", json={"frame": 100})
        try:
            data = resp.json()
            print(f"Raw data: {data}")
            if data["success"]:
                print(f"Seek & Capture success")
                img_data = base64.b64decode(data["image"])
                with open("seek_capture_result.png", "wb") as f:
                    f.write(img_data)
                print("Saved seek_capture_result.png")
            else:
                print(f"Seek & Capture failed: {data.get('error')}")
        except Exception as e:
            print(f"Failed to parse JSON or processing: {e}")
            print(f"Response text: {resp.text}")

async def test_position():
    print("Testing playback position...")
    async with httpx.AsyncClient() as client:
        resp = await client.get(f"{BASE_URL}/preview/position")
        print(f"Position: {resp.json()}")

async def test_record():
    print("Testing audio record (2 seconds)...")
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{BASE_URL}/preview/record", json={"duration_ms": 2000})
        data = resp.json()
        if data["success"]:
            print(f"Record success: {data['bytes_recorded']} bytes, RMS: {data['rms_level']}")
            audio_data = base64.b64decode(data["audio"])
            with open("record_result.wav", "wb") as f:
                f.write(audio_data)
            print("Saved record_result.wav")
        else:
            print(f"Record failed: {data.get('error')}")

async def main():
    await test_capture()
    await test_seek_capture()
    await test_position()
    await test_record()

if __name__ == "__main__":
    asyncio.run(main())
