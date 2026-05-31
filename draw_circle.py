import win32com.client
from PIL import ImageGrab
import time
import os

output_dir = r"C:\Users\王运\Desktop\git\cad_screenshots"
os.makedirs(output_dir, exist_ok=True)

print("=== 步骤1: 启动 AutoCAD 2022 ===")
acad = win32com.client.Dispatch("AutoCAD.Application.24")
acad.Visible = True
time.sleep(3)
print("AutoCAD 已启动并可见")

img = ImageGrab.grab()
img.save(os.path.join(output_dir, "01_autocad_opened.png"))
print("截图已保存: 01_autocad_opened.png")

print()
print("=== 步骤2: 获取模型空间 ===")
doc = acad.ActiveDocument
ms = doc.ModelSpace
print(f"文档: {doc.Name}")

print()
print("=== 步骤3: 画一个圆 (圆心100,100 半径50) ===")
center = [100.0, 100.0, 0.0]
circle = ms.AddCircle(center, 50.0)
print(f"圆已创建! 圆心: ({center[0]}, {center[1]}), 半径: 50")

print()
print("=== 步骤4: 缩放到全图 ===")
acad.ZoomExtents()
time.sleep(1)

img = ImageGrab.grab()
img.save(os.path.join(output_dir, "02_circle_drawn.png"))
print("截图已保存: 02_circle_drawn.png")

print()
print("=== 完成! ===")
print(f"所有截图保存在: {output_dir}")
