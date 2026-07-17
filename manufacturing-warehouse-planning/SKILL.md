---
name: manufacturing-warehouse-planning
description: Generates a multi-SKU manufacturing warehouse planning Excel workbook with inventory calculation formulas for raw material, WIP/process, and finished goods warehouses. Use when a user asks about warehouse inventory sizing or area planning for a factory.
---

# Manufacturing Warehouse Planning Template

This skill generates an Excel workbook for calculating inventory design quantities and estimated floor areas across three warehouse types in a factory. It generates two identical files with different filenames: a standard version and a "七道工序版" version (`仓库规划库存需求计算模板.xlsx` and `仓库规划库存需求计算模板_七道工序版.xlsx`) for business distribution.

## Factory Flow Supported

> 原料 → 粗破 → 烘干 → 磨粉 → 低温碳化 → 石墨化 → 高温碳化 → 成品筛分 → 成品

## Workbook Structure (8 sheets)

| Sheet | Purpose |
|---|---|
| 使用说明 | Chinese instructions for filling the workbook |
| 参数设置 | Global parameters (working days, safety factors, pallet dimensions, utilization rates) |
| 产品SKU | Per-SKU inputs: annual demand, min batch, production interval, target inventory days, buffer hours |
| 原料BOM | Product–raw-material consumption matrix (unit consumption per SKU) |
| 原料仓计算 | Raw-material warehouse: auto-calculates daily consumption, design stock, pallets, area |
| 过程品仓计算 | WIP warehouse per process and SKU: buffer-time WIP, batch WIP, inspection/tail/defect inventory |
| 成品仓计算 | Finished-goods warehouse per SKU: interval/target-days/min-batch stock, safety stock |
| 汇总 | Summary of design stock and estimated area for all three warehouse types |

## Key Formulas

**Raw Material Design Stock:**
```
= MAX(Daily Consumption × Lead Time Days, MOQ) × Safety Factor
  + Safety Stock + Next-batch Preparation Stock
  + Changeover Return Stock + Inspection Pending Stock
```

**WIP Design Stock:**
```
= MAX(Downstream Hourly Consumption × Buffer Hours,
      Batch Size × Max Waiting Batches)
  + Inspection/Waiting Inventory + Tail-batch Inventory
  + Defect/Rework/Isolated Inventory
```

**Finished Goods Design Stock:**
```
= MAX(Daily Demand × Production Interval,
      Daily Demand × Target Inventory Days,
      Minimum Production Batch)
  + Safety Stock (Daily Demand × Safety Days × Safety Factor)
  + Inspection Pending Release Inventory
```

**Area Estimation:**
```
= CEILING(Design Stock / Pallet Load, 1) × Pallet Footprint
  ÷ Stack Layers ÷ Area Utilization Rate × Expansion Factor
```

## How to Generate

```bash
cd manufacturing-warehouse-planning
python generate_warehouse_planning.py
```

This produces `仓库规划库存需求计算模板.xlsx` and an extra same-content copy `仓库规划库存需求计算模板_七道工序版.xlsx` in the same directory.

## Input Conventions

- **Blue cells** (dark blue font, light blue fill): user-editable inputs
- **Green cells** (black font, light green fill): formula-calculated values — do not edit
- The template supports up to **12 SKUs** and **15 raw materials**; extend by copying the last data row's formulas downward
- All monetary units are unspecified — annotate units in the header rows (pieces, kg, boxes, etc.)

## Factors Accounted For

- Multiple products/SKUs with different demand profiles
- Product BOM (raw material consumption per SKU)
- Production changeovers and returned/leftover raw materials after changeovers
- Minimum production batch sizes
- Production intervals/frequencies
- Procurement lead time and MOQ
- Safety stock
- Process buffer inventory between operations
- Batch waiting and process misalignment
- Inspection/quality-hold inventory
- Tail-batch inventory
- Defect/rework/isolated inventory
