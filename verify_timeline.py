import httpx
import asyncio
import json

BASE_URL = "http://localhost:8765/api"

async def get_items():
    async with httpx.AsyncClient() as client:
        response = await client.get(f"{BASE_URL}/items")
        return response.json()["items"]

async def verify():
    items = await get_items()
    print(f"Total items: {len(items)}")
    
    # Check for duplicates
    seen = {}
    duplicates = []
    for item in items:
        if item["type"] in ["VoiceItem", "TextItem"]:
            key = (item["type"], item["text"])
            if key in seen:
                duplicates.append(item)
            else:
                seen[key] = item
    
    if duplicates:
        print(f"FAILED: Found {len(duplicates)} duplicates.")
        for d in duplicates:
            print(f"  {d['type']}: {d['text'][:20]} at Layer {d['layer']}, Frame {d['frame']}")
    else:
        print("SUCCESS: No duplicates found.")
        
    # Check for overlaps
    layers = {}
    for item in items:
        l = item["layer"]
        if l not in layers: layers[l] = []
        layers[l].append(item)
        
    overlaps = []
    for l, l_items in layers.items():
        l_items.sort(key=lambda x: x["frame"])
        for i in range(len(l_items) - 1):
            curr = l_items[i]
            nxt = l_items[i+1]
            if curr["frame"] + curr["length"] > nxt["frame"]:
                overlaps.append((curr, nxt))
                
    if overlaps:
        print(f"FAILED: Found {len(overlaps)} overlaps.")
        for curr, nxt in overlaps:
            print(f"  Layer {curr['layer']}: {curr['text'][:10]}... ends at {curr['frame']+curr['length']}, but next starts at {nxt['frame']}")
    else:
        print("SUCCESS: No overlaps found.")

if __name__ == "__main__":
    asyncio.run(verify())
