# 建議忽略名單（2026-03-21）

## 建議加入 .gitignore

```gitignore
# stitched/輸出資產（圖片 + HTML）
stitch_assets/

# 打包檔
*.zip

# COM 型別庫/二進位介面檔
outlib/
```

## 目前未追蹤項目中，不建議忽略（建議納入版控）

- `ai_roles/`（角色定義文件）
- `app/templates/`（UI 模板）
- `app/ui_server.py`（應用程式碼）
- `run_ui.py`（啟動腳本）
- `ui_tree.py`
- `ui_tree2.py`

## 視用途決定

- `completion_plan.md`：若是個人臨時規劃可忽略；若是團隊共享計畫則建議納入版控。
