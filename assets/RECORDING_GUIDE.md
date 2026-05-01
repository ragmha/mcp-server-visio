# Recording a Demo GIF

## What to Record

Open a terminal and run:

```
copilot

> Create a 3-tier Azure architecture diagram with Front Door,
> VM Scale Sets in 2 availability zones, and Azure SQL with geo-replication
```

Let Copilot call the Excalidraw MCP tools — it will:
1. `create_diagram` → Initializes an empty canvas
2. `add_container` × 3 → Ingress / Compute / Data tier boundaries
3. `add_azure_icon` × 5+ → Front Door, VMSS, SQL, etc.
4. `add_arrow` → Styled connections between services
5. `save_diagram` → Saves as `.excalidraw`
6. `export_diagram` → Exports as PNG

## Recording Tips

- **Window size**: ~1280×720 for good quality
- **Crop**: Just the terminal window
- **Duration**: 30–60 seconds is ideal
- **Tools**: ScreenToGif (Windows), Kap (macOS), Peek (Linux)

## After Recording

1. Save as `assets/demo.gif`
2. Optimize if > 5MB:
   ```bash
   # Using gifsicle (install: npm install -g gifsicle)
   gifsicle -O3 --lossy=80 --resize-width 720 assets/demo.gif -o assets/demo.gif
   ```
3. Commit and push:
   ```bash
   git add assets/demo.gif
   git commit -m "Add demo GIF"
   git push origin main
   ```
