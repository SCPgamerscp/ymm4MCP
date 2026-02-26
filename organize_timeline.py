import httpx
import asyncio
import json

BASE_URL = "http://localhost:8765/api"

async def get_items():
    async with httpx.AsyncClient() as client:
        resp = await client.get(f"{BASE_URL}/items")
        return resp.json()["items"]

async def delete_item(layer, frame):
    async with httpx.AsyncClient() as client:
        resp = await client.post(f"{BASE_URL}/items/delete", json={
            "layer": layer,
            "frame": frame
        })
        return resp.json()

async def organize():
    items = await get_items()
    
    # 1. Deduplication (Same type and text)
    seen = {}
    to_delete = []
    
    for item in items:
        if item["type"] == "VoiceItem" or item["type"] == "TextItem":
            key = (item["type"], item["text"])
            if key in seen:
                # Keep the original one, delete others
                # In this case, maybe keep the one with the earlier frame?
                prev = seen[key]
                if item["frame"] > prev["frame"]:
                    to_delete.append(item)
                else:
                    to_delete.append(prev)
                    seen[key] = item
            else:
                seen[key] = item
    
    print(f"Found {len(to_delete)} duplicates.")
    for item in to_delete:
        print(f"Deleting duplicate: {item['text']} at Layer {item['layer']}, Frame {item['frame']}")
        await delete_item(item["layer"], item["frame"])

    # 2. Overlap fixing (per layer)
    items = await get_items() # Refresh after deletion
    layers = {}
    for item in items:
        l = item["layer"]
        if l not in layers: layers[l] = []
        layers[l].append(item)
    
    for l, l_items in layers.items():
        l_items.sort(key=lambda x: x["frame"])
        for i in range(len(l_items) - 1):
            curr = l_items[i]
            nxt = l_items[i+1]
            if curr["frame"] + curr["length"] > nxt["frame"]:
                overlap = (curr["frame"] + curr["length"]) - nxt["frame"]
                print(f"Overlap detected on Layer {l}: {curr['text'][:10]}... and {nxt['text'][:10]}...")
                # Shift next item
                # we need a set_item_prop or move_item tool
                # The MCP has ymm4_interact(action='edit_item', sub_action='property')
                # Wait, do I have a move command?
                # Yes, in McpHttpServer.cs there is /item/reorder but that's for layer index.
                # Property "Frame" can be set via /item/set_prop
                new_frame = curr["frame"] + curr["length"]
                print(f"Shifting next item to frame {new_frame}")
                async with httpx.AsyncClient() as client:
                    await client.post(f"{BASE_URL}/items/prop", json={
                        "layer": nxt["layer"],
                        "frame": nxt["frame"],
                        "prop": "Frame",
                        "value": str(new_frame)
                    })
                # Update nxt frame for subsequent checks
                nxt["frame"] = new_frame

if __name__ == "__main__":
    asyncio.run(organize())
