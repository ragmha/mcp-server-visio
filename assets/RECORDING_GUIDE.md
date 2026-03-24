# Recording a Demo GIF

## Recommended Tool: ScreenToGif (Windows)

1. **Download**: https://www.screentogif.com/
2. **Install** and open it

## What to Record

Open a terminal and run:

```
copilot

> Create a 3-tier Azure architecture diagram with Front Door, 
> VM Scale Sets in 2 availability zones, and Azure SQL with geo-replication
```

Let Copilot call the Visio MCP tools — it will:
1. `create_diagram` → Visio opens with a blank landscape page
2. `add_tier_band` × 3 → Ingress / Compute / Data tiers
3. `add_azure_shape` × 5+ → Front Door, VMSS, SQL, etc.
4. `add_container` → Availability zone boundaries
5. `connect_shapes` → Styled connectors between services
6. `export_page` → Saves as PNG

## Recording Tips

- **Window size**: ~1280×720 for good quality
- **Crop**: Just the terminal + Visio window side by side
- **Duration**: 30–60 seconds is ideal
- **Frame rate**: 15 FPS keeps the file small

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
