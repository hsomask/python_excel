# import re
# import os
#
# def main():
#     # 指定输入和输出文件路径
#     input_file_path = 'C:\\Users\\hsoluo\\Desktop\\testDemo.txt'  # 请将此路径替换为实际的文件路径
#     output_file_path = 'C:\\Users\\hsoluo\\Desktop\\123.txt'
#
#     # 打开输入文件并读取内容
#     with open(input_file_path, 'r', encoding='utf-8') as infile:
#         content = infile.read()
#
#    # 使用正则表达式查找所有数字，并在其周围添加单引号和逗号
#     # 首先添加单引号，并在数字后面添加逗号
#     modified_content = re.sub(r'(\d+(\.\d+)?)', r"'\1',", content)
#
#     # 移除最后一行的逗号
#     lines = modified_content.split('\n')
#     if lines[-1].endswith(','):
#         lines[-1] = lines[-1][:-1]
#     modified_content = '\n'.join(lines)
#
#     # 将修改后的内容写入输出文件
#     try:
#         # 确保输出目录存在
#         os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
#
#         with open(output_file_path, 'w', encoding='utf-8') as outfile:
#             outfile.write(modified_content)
#
#         print(f"处理完成，结果已保存到 {output_file_path}")
#     except Exception as e:
#         print(f"保存文件时发生错误: {str(e)}")
#
#
#
# if __name__ == "__main__":
#     main()
