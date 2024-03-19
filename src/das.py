import os
import time

import gradio as gr
import tempfile
import shutil
from excel_process import excel_process

def check_path(path):
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"create {path}")
    print(f"{path} exists")

def generate_file(files_obj, invoice):
    global tmpdir
    print('临时文件夹地址：{}'.format(tmpdir))
    print('上传文件的地址：{}'.format(files_obj)) # 输出上传后的文件在gradio中保存的绝对地址
    print(f'发票号：{invoice}')

    #获取到上传后的文件的绝对路径后，其余的操作就和平常一致了

    # 将文件复制到临时目录中
    for file_obj in files_obj:
        shutil.copy(file_obj, tmpdir)

    # 获取上传Gradio的文件名称
    FilesName = []
    for file_obj in files_obj:
        FileName=os.path.basename(file_obj.name)
        FilesName.append(FileName)
    output_path = "./output/"
    check_path(output_path)
    FileName = f'Receive Upload_BL No.COSU{int(time.time())}.xlsx'
    OutputFilePath=os.path.join(output_path, FileName)
    # 获取拷贝在临时目录的新的文件地址
    for FileName in FilesName:
        InputFilePath=os.path.join(tmpdir, FileName)
        excel_process(InputFilePath, OutputFilePath, invoice)

    # # del tmpdir
    # try:
    #     shutil.rmtree(tmpdir)
    #     print(f'Folder {tmpdir} and its content removed')
    # except:
    #     print(f'Folder {tmpdir} not deleted')
    # 返回新文件的的地址（注意这里）
    return OutputFilePath

def demo():
    global tmpdir
    check_path('./tmpdir')
    with tempfile.TemporaryDirectory(dir='./tmpdir') as tmpdir:
        # 定义输入和输出
        inputs = gr.components.File(label="上传文件",file_count="multiple") ## 支持多文件上传
        invoice = gr.components.Textbox(label='发票号', lines=1)
        outputs = gr.components.File(label="下载文件")

        # 创建 Gradio 应用程序g
        app = gr.Interface(fn=generate_file, inputs=[inputs,invoice], outputs=outputs,   title="DAS Solar数据处理模块",
                      description="上传.xlsx文件,每个文件处理时间约3s,遇到问题请联系Eric Yang")
        # 启动应用程序
        app.launch(share=True,inbrowser=True, server_name ='0.0.0.0')
        # app.launch(inbrowser=True, server_name ='0.0.0.0')

if __name__=="__main__":
    demo()
