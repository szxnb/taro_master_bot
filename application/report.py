from datetime import datetime
from pptx import Presentation


def replace_content(shape, target_text, replacement_text):
    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if target_text in run.text:
                    run.text = run.text.replace(target_text, replacement_text)


def generate_ppt(part1, part2, part3, part4):
    presentation = Presentation("../templates/template.pptx")
    for slide_number, slide in enumerate(presentation.slides, start=1):
        for shape in slide.shapes:
            replace_content(shape, "{{#1overview}}", part1)
            replace_content(shape, "{{#2love}}", part2)
            replace_content(shape, "{{#3career}}", part3)
            replace_content(shape, "{{#4finance}}", part4)
    # 格式化当前时间
    formatted_time = datetime.now().strftime("%Y%m%d%H%M%S")
    report_name = "reports/TaroReport_" + "3500466989@qq.com_" + formatted_time + ".pptx"
    print(report_name)
    presentation.save("../" + report_name)
    print("生成PPT完成!")