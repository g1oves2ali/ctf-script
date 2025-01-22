import os
import email
from email import policy
from email.parser import BytesParser

def extract_attachments_from_eml(eml_folder, output_folder):
    """
    从指定文件夹中的所有 .eml 文件提取附件并保存。
    
    参数：
        eml_folder (str): 包含 .eml 文件的文件夹路径。
        output_folder (str): 附件保存的目标文件夹路径。
    """
    # 确保输出目录存在
    os.makedirs(output_folder, exist_ok=True)
    
    # 遍历文件夹中的所有 .eml 文件
    for filename in os.listdir(eml_folder):
        if filename.lower().endswith('.eml'):
            eml_path = os.path.join(eml_folder, filename)
            print(f"正在处理文件: {eml_path}")

            try:
                # 打开并解析 .eml 文件
                with open(eml_path, 'rb') as eml_file:
                    msg = BytesParser(policy=policy.default).parse(eml_file)
                
                # 遍历邮件的所有附件
                for part in msg.iter_attachments():
                    # 获取附件的文件名（确保兼容中文）
                    attachment_filename = part.get_filename()
                    if attachment_filename:
                        # 确保文件名的合法性
                        safe_filename = "".join(c if c.isalnum() or c in "._-" else "_" for c in attachment_filename)
                        attachment_path = os.path.join(output_folder, safe_filename)

                        # 提取附件内容
                        with open(attachment_path, 'wb') as attachment_file:
                            attachment_file.write(part.get_payload(decode=True))
                        
                        print(f"附件已提取: {attachment_path}")
            except Exception as e:
                print(f"处理文件 {eml_path} 时出错: {e}")

if __name__ == "__main__":
    # 输入文件夹路径（包含 .eml 文件）
    eml_folder = input("请输入包含 .eml 文件的文件夹路径: ").strip()

    # 输出附件保存路径
    output_folder = input("请输入附件保存的文件夹路径: ").strip()

    # 提取附件
    extract_attachments_from_eml(eml_folder, output_folder)
    print("附件提取完成。")
