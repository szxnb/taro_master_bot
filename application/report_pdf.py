from docx import Document
from docx.shared import Pt, RGBColor
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches

user_question = '11Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?'
master_words = '22Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?Can I get married this year?'
# Name
card_first_name = 'CardⅠ The MAGIGIAN'
card_second_name = 'CardⅡ The MAGIGIAN'
card_third_name = 'CardⅢ The MAGIGIAN'
# Image
card_first_image = '../assets/love.png'
card_second_image = '../assets/love.png'
card_third_image = '../assets/love.png'
# Description
card_first_description = '111The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '
card_second_description = '222The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '
card_third_description = '333The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '

summarize = '333The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '
love_suggestion = '333The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '
career_suggestion = '333The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '
wealth_suggestion = '333The Magician is one tarot card that is filled with symbolism. The central figure depicts someone with one hand pointed to the sky, while the other hand points to the ground, as if to say "as above, so below". This is a rather complicated phrase, but its summarization is that earth reflects heaven, the outer world reflects within, the microcosm reflects the macrocosm, earth reflects God. It can also be interpreted here that the magician symbolizes the ability to act as a go-between between the world above and the contemporary, human world. '

# 读取文档
doc = Document('../templates/template.docx')

def limit_run_length(run, max_length):
    if len(run.text) > max_length:
        run.text = run.text[:max_length - 3] + '...'

# 遍历文档中的所有表格
for table in doc.tables:
    # 遍历表格中的所有行
    for row in table.rows:
        # 遍历行中的所有单元格
        for cell in row.cells:
            # 遍历单元格中的所有段落
            for paragraph in cell.paragraphs:
                if '{{#date}}' in paragraph.text:
                    # 设置段落的对齐方式
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    # 遍历段落中的所有运行
                    for run in paragraph.runs:
                        # 替换运行中的文字
                        run.text = run.text.replace('{{#date}}', 'October 10,2024')
                        # 限制运行的长度
                        limit_run_length(run, 300)
                if '{{#user_question}}' in paragraph.text:
                    # 设置段落的对齐方式
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    # 设置段落的行间距
                    # 遍历段落中的所有运行
                    for run in paragraph.runs:
                        # 替换运行中的文字
                        run.text = run.text.replace('{{#user_question}}', user_question)
                        # 限制运行的长度
                        limit_run_length(run, 300)
                if '{{#master_words}}' in paragraph.text:
                    # 设置段落的对齐方式
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    # 遍历段落中的所有运行
                    for run in paragraph.runs:
                        # 替换运行中的文字
                        run.text = run.text.replace('{{#master_words}}', master_words)
                        # 限制运行的长度
                        limit_run_length(run, 300)
                if '{{#card_first_name}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_first_name}}', card_first_name)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_first_description}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_first_description}}', card_first_description)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_second_name}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_second_name}}', card_second_name)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_second_description}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_second_description}}', card_second_description)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_third_name}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_third_name}}', card_third_name)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_third_description}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#card_third_description}}', card_third_description)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#summarize}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#summarize}}', summarize)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#love_suggestion}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#love_suggestion}}', love_suggestion)
                        # 限制运行的长度
                        limit_run_length(run, 300)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#career_suggestion}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#career_suggestion}}', career_suggestion)
                        # 限制运行的长度
                        limit_run_length(run, 300)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#wealth_suggestion}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        paragraph.text = paragraph.text.replace('{{#wealth_suggestion}}', wealth_suggestion)
                        # 限制运行的长度
                        limit_run_length(run, 300)
                        paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_first_image}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        # 清空单元格中的文本
                        cell.text = ''
                        # 向单元格中插入图片
                        cell.paragraphs[0].add_run().add_picture('../assets/Card_1_Fool.png', width = Inches(2), height = Inches(3))
                        # paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_second_image}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        # 清空单元格中的文本
                        cell.text = ''
                        # 向单元格中插入图片
                        cell.paragraphs[0].add_run().add_picture('../assets/Card_1_Fool.png', width = Inches(2), height = Inches(3))
                        # paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                if '{{#card_third_image}}' in paragraph.text:
                        # 设置段落的对齐方式
                        paragraph.style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        original_style = paragraph.style
                        original_runs = paragraph.runs
                        # 清空单元格中的文本
                        cell.text = ''
                        # 向单元格中插入图片
                        cell.paragraphs[0].add_run().add_picture('../assets/Card_1_Fool.png', width = Inches(2), height = Inches(3))
                        # paragraph.style = original_style
                        for run in paragraph.runs:
                            run.bold = original_runs[0].bold
                            run.italic = original_runs[0].italic
                            run.underline = original_runs[0].underline
                            run.font.size = original_runs[0].font.size
                            run.font.name = original_runs[0].font.name
                            run.font.color.rgb = original_runs[0].font.color.rgb
                
 # 格式化当前时间
formatted_time = datetime.now().strftime("%Y%m%d%H%M%S")
report_name = "reports/TarotReport_" + "3500466989@qq.com_" + formatted_time + ".docx"
print("生成完成!")
# 保存修改后的文档
doc.save("../" + report_name)


