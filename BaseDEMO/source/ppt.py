# from pptx import Presentation
#
# # 创建新PPT
# # prs = Presentation()
#
# # 打开现有PPT
# prs = Presentation("Z:\\1.pptx")
#
# slide = prs.slides[0]  # 获取第一张幻灯片
#
# slide_layout = prs.slide_layouts[1]  # 使用第二个布局（标题+内容）
# new_slide = prs.slides.add_slide(slide_layout)
#
# from pptx.util import Inches
#
# # 在指定位置添加文本框
# textbox = slide.shapes.add_textbox(left=Inches(1), top=Inches(1),
#                                    width=Inches(4), height=Inches(2))
# text_frame = textbox.text_frame
# text_frame.text = "自定义文本"
#
# slide.shapes.add_picture(r"Z:\1.png",
#                          left=Inches(2), top=Inches(3),
#                          width=Inches(5), height=Inches(3))
#
# from pptx.enum.text import PP_ALIGN
#
# table = slide.shapes.add_table(rows=3, cols=3,
#                               left=Inches(1), top=Inches(4),
#                               width=Inches(5), height=Inches(2)).table
# # 填充表格内容
# for row in table.rows:
#     for cell in row.cells:
#         cell.text = "单元格"
#         cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
#
# from pptx.chart.data import CategoryChartData
# from pptx.enum.chart import XL_CHART_TYPE
#
# # 准备数据
# from pptx import Presentation
# from pptx.chart.data import CategoryChartData
# from pptx.enum.chart import XL_CHART_TYPE
# from pptx.util import Inches
#
# slide = prs.slides.add_slide(prs.slide_layouts[5])
# chart_data = CategoryChartData()
# chart_data.categories = ["Q1", "Q2", "Q3"]
# chart_data.add_series("Sales", (19.2, 21.4, 16.7))
#
# # 定义坐标和尺寸（正确参数名）
# x, y, cx, cy = Inches(1), Inches(1), Inches(6), Inches(4)
#
# # 添加图表（按位置传参）
# chart = slide.shapes.add_chart(
#     XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data
# )
# # 保存结果
# prs.save(r"Z:\1.pptx")

#==========================================================================

# from pptx import Presentation
# from pptx.util import Cm, Pt
# from pptx.enum.text import PP_ALIGN
# from pptx.dml.color import RGBColor
# from pptx.chart.data import CategoryChartData
# from pptx.enum.chart import XL_CHART_TYPE
#
#
# def create_ppt_slide(output_path):
#     prs = Presentation()
#     # 改用标题版式（索引0为标题版式）
#     slide = prs.slides.add_slide(prs.slide_layouts[0])
#
#     # 设置背景色（保持原图深蓝渐变）
#     background = slide.background
#     fill = background.fill
#     fill.gradient()
#     fill.gradient_stops[0].color.rgb = RGBColor(0, 32, 96)
#     fill.gradient_stops[1].color.rgb = RGBColor(0, 64, 128)
#
#     # 获取标题占位符并设置文本
#     title_shape = slide.shapes.title
#     title_shape.text = "监理项目部计划到位情况"
#
#     # 保持原图标题样式
#     title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)
#     title_shape.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
#
#     # 添加说明文字（居中白底蓝框）
#     textbox = slide.shapes.add_textbox(Cm(2), Cm(3), Cm(20), Cm(4))
#     text_frame = textbox.text_frame
#     text_frame.word_wrap = True
#     text_frame.paragraphs[0].text = "2月20日公司监理业务计划开工项目228项...（具体文字内容）"
#     text_frame.paragraphs[0].font.size = Pt(14)
#
#     # 文字框样式设置
#     textbox.fill.solid()
#     textbox.fill.fore_color.rgb = RGBColor(240, 240, 240)  # 浅灰背景
#     textbox.line.color.rgb = RGBColor(0, 112, 192)  # 蓝色边框
#     textbox.line.width = Pt(1.5)
#
#     # 计算居中位置
#     textbox.left = int((prs.slide_width - textbox.width) / 2)
#     textbox.top = title_shape.top + title_shape.height + Cm(1).emu
#
#     # 配置图表数据
#     chart_data = CategoryChartData()
#     chart_data.categories = ['项目总数', '中高风险', '总监上报', '续任计划']
#     chart_data.add_series('计划到位', (193, 37, 230, 231))
#     chart_data.add_series('实际到位', (228, 37, 193, 231))
#
#     # 添加图表（自动居中）
#     chart_shape = slide.shapes.add_chart(
#         XL_CHART_TYPE.COLUMN_CLUSTERED,
#         Cm(2), Cm(8), Cm(20), Cm(10),
#         chart_data
#     )
#     chart = chart_shape.chart
#
#     # 图表样式设置（保持原图蓝橙配色）
#     plot = chart.plots[0]
#     plot.has_data_labels = True
#     series_1, series_2 = plot.series[0], plot.series[1]
#
#     # 蓝色系列（对应原图左侧柱）
#     series_1.format.fill.solid()
#     series_1.format.fill.fore_color.rgb = RGBColor(0, 112, 192)
#
#     # 橙色系列（对应原图右侧柱）
#     series_2.format.fill.solid()
#     series_2.format.fill.fore_color.rgb = RGBColor(237, 125, 49)
#
#     # 自动居中（完美还原原图布局）
#     chart_shape.left = int((prs.slide_width - chart_shape.width) / 2)
#
#     prs.save(output_path)
#
# # 使用示例
# create_ppt_slide(r"Z:\监理数据报告.pptx")

#=====================================================================================

from pptx.dml.color import RGBColor
from pptx.util import Cm, Pt
from pptx import Presentation
import os

def insert_image_to_ppt(ppt_path, slide_index, img_path, position='center'):
    """ 企业级PPT图片插入方案（严格匹配原图规范） """
    # 验证文件可达性
    if not os.access(img_path, os.R_OK):
        raise FileNotFoundError(f"图片路径不可访问: {img_path}")

    prs = Presentation(ppt_path) if os.path.exists(ppt_path) else Presentation()

    # 智能版式管理系统
    while len(prs.slides) <= slide_index:
        new_slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
        # 精准还原原图渐变背景
        background = new_slide.background
        fill = background.fill
        fill.gradient()
        fill.gradient_stops[0].color.rgb = RGBColor(0, 32, 96)  # 深蓝
        fill.gradient_stops[1].color.rgb = RGBColor(0, 64, 128)  # 渐变蓝

    slide = prs.slides[slide_index]

    # 专业定位算法（匹配原图布局）
    slide_width = prs.slide_width
    base_top = Cm(3.8)  # 精确测量原图标题下沿位置

    position_config = {
        'center': (int((slide_width - Cm(15)) / 2), base_top),
        'left': (Cm(2.5), base_top)  # 左侧边距微调
    }

    try:
        left, top = position_config[position.lower()]
    except KeyError:
        raise ValueError("定位参数仅支持 'center' 或 'left'")

    # 插入图片（保持原图比例）
    try:
        pic = slide.shapes.add_picture(img_path, left, top, width=Cm(20))
        # 应用原图样式规范
        pic.line.color.rgb = RGBColor(0, 112, 192)  # 主题蓝边框
        pic.line.width = Pt(1.2)  # 精确到0.1pt
    except Exception as e:
        raise RuntimeError(f"图形引擎错误: {str(e)}")

    prs.save(ppt_path)


# 企业级调用示例
try:
    # 主视觉插入（居中定位）
    insert_image_to_ppt(
        ppt_path=r"Z:\监理数据报告.pptx",
        slide_index=0,
        img_path=r"Z:\temp_0\2025-03-16_050735.png",
        position='center'
    )

    # 辅助说明插入（左侧对齐）
    insert_image_to_ppt(
        ppt_path=r"Z:\监理数据报告.pptx",
        slide_index=1,
        img_path=r"Z:\temp_0\2025-03-16_050049.png",
        position='left'
    )
except Exception as e:
    print(f"企业级操作失败: {e}")