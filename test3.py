import os
from difflib import unified_diff
import docx

def doc_to_text(doc_path):
    # 使用 python-docx 库读取 Word 文档内容
    doc = docx.Document(doc_path)
    text = '\n'.join(paragraph.text for paragraph in doc.paragraphs)
    return text

def compare_docs(doc1_path, doc2_path):
    # 将 Word 文档转换为文本
    doc1_text = doc_to_text(doc1_path)
    doc2_text = doc_to_text(doc2_path)

    # 使用 difflib 比较文本内容
    diff = unified_diff(doc1_text.splitlines(), doc2_text.splitlines(),
                         fromfile='Document 1', tofile='Document 2', lineterm='')
    #for line in diff:
    #    print(line)
    diff_text = [line for line in diff if not line.startswith('+++') and not line.startswith('---') and (line.startswith('-') or line.startswith('+') )]

    identical_lines = len(diff_text)
    total_lines = len(doc1_text.splitlines()) + len(doc2_text.splitlines())

    if total_lines == 0:
        return 1.0 if identical_lines == 0 else 0.0

    duplication_rate = 1 - identical_lines / total_lines
    return duplication_rate

def compare_all_docs(docs_list):
    similarities = []
    for i in range(len(docs_list)):
        for j in range(i + 1, len(docs_list)):
            similarity = compare_docs(os.path.join(base_path,docs_list[i]), os.path.join(base_path,docs_list[j]))
            similarities.append((docs_list[i], docs_list[j], similarity))
            #print("{}-{}:{}".format(similarity,docs_list[i],docs_list[j]))
    return similarities
def list_to_txt(lst, filename):
    with open(filename, 'w') as f:
        for item in lst:
            f.write(str(item) + '\n')

base_path = r"C:\Users\HuangMin\Desktop\temp"
docs_list = os.listdir(base_path)
results = compare_all_docs(docs_list)
list_to_txt(results,'results.txt')
for doc1, doc2, similarity in results:
    print(f"{os.path.basename(doc1)} and {os.path.basename(doc2)} are {similarity:.2f}% similar.")