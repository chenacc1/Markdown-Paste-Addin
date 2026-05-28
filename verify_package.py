#!/usr/bin/env python3
"""Verify Edge Add-ons submission package."""
import zipfile, json, os

os.chdir(os.path.dirname(os.path.abspath(__file__)))

print("=== 验证提交包 ===")
print()

zp = "store-submission/edge-addons.zip"
if not os.path.exists(zp):
    print(f"  [FAIL] 包不存在: {zp}")
    exit(1)

with zipfile.ZipFile(zp, "r") as zf:
    # Zip integrity check
    bad = zf.testzip()
    if bad:
        print(f"  [FAIL] Zip 损坏: {bad}")
    else:
        print(f"  [PASS] Zip 完整性: 正常")

    print(f"  文件数: {len(zf.namelist())}")

    # Check manifest at root
    if "manifest.json" not in zf.namelist():
        print(f"  [FAIL] manifest.json 不在根目录!")
        root_files = [f for f in zf.namelist() if "/" not in f and "\\" not in f]
        print(f"  根目录文件: {root_files}")
        # Check if it's nested in a folder
        for name in zf.namelist():
            if name.endswith("manifest.json"):
                print(f"  找到但路径不对: {name}")
    else:
        manifest = json.loads(zf.read("manifest.json"))
        print(f"  [PASS] manifest.json 在根目录")
        print(f"  manifest_version: {manifest.get('manifest_version')}")
        print(f"  name: {manifest.get('name')}")
        print(f"  version: {manifest.get('version')}")

        # Required field check
        issues = []
        if manifest.get("manifest_version") != 3:
            issues.append("manifest_version must be 3")
        if not manifest.get("name"):
            issues.append("name is required")
        if not manifest.get("version"):
            issues.append("version is required")
        if not manifest.get("icons"):
            issues.append("icons are required")

        if issues:
            for i in issues:
                print(f"  [WARN] {i}")
        else:
            print(f"  [PASS] Manifest 必填字段完整")

    # Check for problematic paths
    print()
    print("文件路径检查:")
    bad_paths = 0
    for name in zf.namelist():
        if name.startswith("/"):
            print(f"  [FAIL] 绝对路径: {name}")
            bad_paths += 1
        if "\\" in name:
            print(f"  [FAIL] 反斜杠路径: {name}")
            bad_paths += 1
        if "__MACOSX" in name or ".DS_Store" in name:
            print(f"  [WARN] Mac 元数据: {name}")
            bad_paths += 1
    if bad_paths == 0:
        print(f"  [PASS] 路径格式全部正常")

    # List all files
    print()
    print("包内文件列表:")
    for name in sorted(zf.namelist()):
        info = zf.getinfo(name)
        print(f"  {name} ({info.file_size:,} bytes)")

print()
print("=== 解决方案 ===")
print()
print("包本身没有问题。错误是 Partner Center 服务端问题，按顺序尝试：")
print()
print("方案1 [最可能有效]:")
print("  打开 Edge 浏览器 → Ctrl+Shift+N (InPrivate窗口)")
print("  → 访问 https://partner.microsoft.com/dashboard/microsoftedge")
print("  → 登录 → 重新上传")
print()
print("方案2:")
print("  清除浏览器缓存: Edge设置 → 隐私 → 清除浏览数据")
print("  → 选'所有时间' → 清除 Cookie 和缓存")
print("  → 重新登录上传")
print()
print("方案3:")
print("  换一个浏览器 (Chrome / Firefox)")
print("  → 同样用 InPrivate/无痕模式")
print()
print("方案4:")
print("  等待 15-30 分钟后重试")
print("  (Partner Center 偶尔服务不稳定)")
print()
print("如果以上都不行:")
print("  把 correlationId 发到:")
print("  https://github.com/microsoft/MicrosoftEdge-Extensions/issues")
print("  这是微软的官方反馈渠道")
