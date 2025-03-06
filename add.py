import win32com.client as win32
from collections import OrderedDict

def process_word_document(file_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = True  # 保持可见以观察执行过程

    try:
        doc = word.Documents.Open(file_path)
        unique_items = OrderedDict()

        # 创建独立的查找范围
        search_range = doc.Content
        search_range.Find.ClearFormatting()

        # 使用更可靠的通配符模式
        finder = search_range.Find
        finder.Text = "【[!】]{1,}】"  # 安全匹配模式
        finder.MatchWildcards = True
        finder.Forward = True
        finder.Wrap = win32.constants.wdFindStop

        # 安全查找逻辑
        while True:
            found = finder.Execute()
            if not found:
                break
            
            # 精确提取匹配内容
            if search_range.Text.startswith("【") and search_range.Text.endswith("】"):
                clean_text = search_range.Text[1:-1].strip()
                unique_items[clean_text] = None

            # 移动查找范围到当前匹配之后
            search_range.SetRange(max(search_range.End, search_range.Start + 1), doc.Content.End)

        if unique_items:
            # 在文档末尾创建列表
            end_range = doc.Content
            end_range.Collapse(win32.constants.wdCollapseEnd)
            end_range.InsertParagraphAfter()
            
            # 添加列表标题
            title_range = end_range.Paragraphs.Add().Range
            title_range.Text = "匹配项列表：\n"
            
            # 插入列表项
            list_range = title_range.Document.Range(title_range.End, title_range.End)
            for item in unique_items.keys():
                list_range.InsertAfter(f"{item}\n")
            
            # 应用列表格式（自动编号）
            list_range = title_range.Document.Range(title_range.End, list_range.End - 1)
            list_range.ListFormat.ApplyNumberDefault()

        doc.Save()
        print("处理成功，已保存文档")

    except Exception as e:
        print(f"执行出错：{str(e)}")
    finally:
        doc.Close(SaveChanges=win32.constants.wdDoNotSaveChanges)
        word.Quit()

if __name__ == "__main__":
    file_path = r"D:\workSpace\group9\shenbao\example.docx"  # 修改为实际路径
    process_word_document(file_path)