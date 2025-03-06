import win32com.client as win32
from collections import OrderedDict

def replace_cross_references(file_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False  # 生产环境建议保持False
    doc = word.Documents.Open(file_path)
    wdconst = win32.constants

    try:
        # 收集自动编号项
        numbered_items = OrderedDict()
        item_index = 0
        
        # 第一遍扫描：识别所有自动编号项
        for para in doc.Paragraphs:
            rng = para.Range
            if rng.ListFormat.ListType != wdconst.wdListNoNumbering:
                # 提取纯净文本内容（去除换行符和空白）
                clean_text = rng.Text.replace('\r', '').replace('\x07', '').strip()
                if clean_text:
                    item_index += 1
                    numbered_items[clean_text] = item_index
        #输出编号项
        print(numbered_items)

        # 第二遍处理：查找替换【】内容
        find = doc.Content.Find
        find.ClearFormatting()
        find.Text = "【*】"
        find.MatchWildcards = True
        find.Forward = True
        find.Wrap = wdconst.wdFindStop

        while find.Execute():
            found_range = find.Parent.Duplicate
            original_text = found_range.Text[1:-1].strip()  # 去除【】符号
            
            if original_text in numbered_items:
                # 清空原始内容区域
                found_range.Text = ""
                
                # 修改点：InsertAsHyperlink=True 添加超链接功能
                found_range.InsertCrossReference(
                    ReferenceType=wdconst.wdRefTypeNumberedItem,
                    ReferenceKind=wdconst.wdNumberFullContext,
                    ReferenceItem=numbered_items[original_text],
                    InsertAsHyperlink=True,  # ← 这里改为True
                    IncludePosition=False
                )

                # 格式修正（可选）
                found_range.Font.Reset()

        doc.Save()
        print("成功更新所有交叉引用！")

    except Exception as e:
        print(f"操作失败，原因：{str(e)}")
        print("建议检查：1. Word自动编号是否规范 2. 文档是否被锁定 3. 特殊符号使用")
    finally:
        doc.Close(SaveChanges=True)
        word.Quit()

# 使用示例（注意转义路径）
replace_cross_references(r"D:\workSpace\group9\shenbao\example.docx")