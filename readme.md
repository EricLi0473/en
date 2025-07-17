# 英语单词默写与错题本 PDF 生成工具

本软件是一个用于英语单词默写练习和错题本管理的命令行工具。它支持从 Excel 文件批量导入单词，通过多种模式（全默写、抽查、错题本）自动生成默写题目和答案的 PDF 文件，方便教师或学生打印练习。同时，软件还支持错题本的添加、删除和单独输出，帮助用户针对易错单词进行针对性复习。所有功能均可通过命令行参数灵活调用。

## 使用步骤

1. 确保你的excel表格符合以下列名
   [序号]  *格式必须为‘’List数字-数字‘’形式
   [中文]
   [词性]
   [英文]
   *我提供了英文单词excel为模板
2. 安装requirement.txt 内的库
3. 使用命令行启动
   python dictation_gui.py

## 命令案例

```
📘 1. 全默写指定 List（例：List 1, 5, 7）
python dictation.py --excel 英文单词.xlsx generate --mode full --lists 1,5,7

🎲 2. 抽查模式（例：从 List 2, 3 随机抽取 30 个）
python dictation.py --excel 英文单词.xlsx generate --mode sample --lists 2,3 --count 30

📘 3. 加入错题本内容一起输出（例：List 1 + 错题本）
python dictation.py --excel 英文单词.xlsx generate --mode full --lists 1 --include-wb

✍️ 4. 添加错题本条目（交互式输入 10-1、2-5 等）
python dictation.py --excel 英文单词.xlsx wb add

❌ 5. 从错题本中删除条目
python dictation.py --excel 英文单词.xlsx wb remove

🧾 6. 单独输出错题本为 PDF（题目 + 答案）
python dictation.py --excel 英文单词.xlsx wb output

🌱 7. 设置随机种子（可复现的抽样）
python dictation.py --excel 英文单词.xlsx generate --mode sample --lists 1,2,3 --count 20 --seed 42
```


