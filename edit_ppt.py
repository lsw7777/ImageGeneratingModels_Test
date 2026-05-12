import pptx
from pptx.util import Pt, Inches, Emu
import copy
import sys
sys.stdout.reconfigure(encoding='utf-8')

prs = pptx.Presentation('国产图片生成模型测试分析报告.pptx')

# ============================================================
# Slide 1 - Cover
# ============================================================
slide1 = prs.slides[0]

# Shape[1]: 万相 / 即梦 / 文心 / 混元 / GLM → 万相 / 即梦 / 文心 / 混元 / GLM / Z-Image
shape1_s1 = slide1.shapes[1]
shape1_s1.text = '万相 / 即梦 / 文心 / 混元 / GLM / Z-Image'

# Shape[3]: 6 类场景 · 18 个测试用例 · 98 张候选图片 → 112张
shape3_s1 = slide1.shapes[3]
shape3_s1.text = '6 类场景 · 18 个测试用例 · 112 张候选图片'

# Shape[7]: • 评价指标... • 六类工作场景... • 首选... → 加入Z-Image
shape7_s1 = slide1.shapes[7]
shape7_s1.text = '• 评价指标：成本、质量、信息密度、提示词贴合、文字控制\n• 六类工作场景下的典型测试用例对比\n• 首选：万相；备选：即梦、Z-Image、GLM'

# Shape[14]: GLM → 新增 Z-Image 和 GLM 卡片
# Shape[14]: "GLM" → "Z-Image"
shape14_s1 = slide1.shapes[14]
shape14_s1.text = 'Z-Image'

# Shape[15]: "图标简洁、文案可控的基础素材" → Z-Image 描述
shape15_s1 = slide1.shapes[15]
shape15_s1.text = '长文本、Logo 与表格场景表现突出'

# Add GLM as a new card at bottom
# Copy shapes 14-15 and reposition them as GLM
from lxml import etree

def clone_shape_and_set(text, source_shape, slide, top_offset_emu):
    """Clone a text box shape and set its text"""
    # Clone the XML element
    sp = copy.deepcopy(source_shape._element)
    # Update text
    txBody = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}txBody')
    if txBody is not None:
        for rg in txBody.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}r'):
            t_elem = rg.find('{http://schemas.openxmlformats.org/drawingml/2006/main}t')
            if t_elem is not None:
                t_elem.text = text
    # Update vertical position
    spPr = sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}spPr')
    if spPr is not None:
        xfrm = spPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
        if xfrm is not None:
            off = xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off')
            if off is not None:
                off.set('y', str(top_offset_emu))
    slide.shapes._spTree.append(sp)
    return slide.shapes[-1]

# Add GLM text box below Z-Image
# Original Shape[14] is at some y position, we add new shapes below
# Find the y position of shape[14]
s14_sp = shape14_s1._element
s14_xfrm = s14_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
s14_y = int(s14_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))

# Get height of shape[15] to calculate gap
s15_sp = shape15_s1._element
s15_xfrm = s15_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
s15_y = int(s15_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))
s15_h = int(s15_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext').get('cy'))

# New GLM position: below Z-Image description
new_glm_y = s15_y + s15_h + Inches(0.3)
new_glm_desc_y = new_glm_y + Inches(0.45)

clone_shape_and_set('GLM', shape14_s1, slide1, new_glm_y)
clone_shape_and_set('图标简洁、中文标签可控的基础素材', shape15_s1, slide1, new_glm_desc_y)

# ============================================================
# Slide 2 - 综合结论
# ============================================================
slide2 = prs.slides[1]

# Shape[12]: "备选 2" → "备选 2"
# Shape[13]: "GLM" → "Z-Image"
# Shape[14]: description → Z-Image description
shape13_s2 = slide2.shapes[13]
shape13_s2.text = 'Z-Image'
shape14_s2 = slide2.shapes[14]
shape14_s2.text = '长文本、Logo 与表格场景表现最佳，适合品牌视觉和结构化图表'

# Add new "备选 3" section for GLM below 备选 2
s12_sp = shape13_s2._element
s12_xfrm = s12_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
s12_y = int(s12_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))
s12_h = int(s12_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext').get('cy'))
s14_y2 = int(s15_s1._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm').find(
    '{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))

# Find the y position of shape[14] (which is now Z-Image description)
shape14_s2_sp = shape14_s2._element
shape14_s2_xfrm = shape14_s2_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
shape14_desc_y = int(shape14_s2_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))
shape14_desc_h = int(shape14_s2_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext').get('cy'))

new_glm_label_y = shape14_desc_y + shape14_desc_h + Inches(0.3)
new_glm_text_y = new_glm_label_y + Inches(0.4)

# Create GLM label (备选 3)
shape12_s2 = slide2.shapes[12]  # "备选 2" label
clone_shape_and_set('备选 3', shape12_s2, slide2, new_glm_label_y)
clone_shape_and_set('GLM', slide2.shapes[13], slide2, new_glm_text_y)

# Find y for GLM description
glm_text_sp = slide2.shapes[-1]._element
glm_text_xfrm = glm_text_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
glm_text_y = int(glm_text_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))
glm_text_h = int(glm_text_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ext').get('cy'))
glm_desc_y = glm_text_y + glm_text_h + Inches(0.05)
clone_shape_and_set('干净克制，适合 PPT 背景、图标底图和中文标签可控的基础素材。',
                    shape14_s2, slide2, glm_desc_y)

# Shape[17]: 核心判断 → 更新
shape17_s2 = slide2.shapes[17]
shape17_s2.text = ('• 万相综合领先，Z-Image在长文本和Logo场景表现突出；万相强在画面和信息密度，'
                   '即梦强在成本、中文可读性与批量稳定。\n'
                   '• 文心和混元在本次样本中提示词理解与文字语言一致性问题较明显，不建议作为主生产链路。\n'
                   '• 所有模型生成的图表只适合作为概念插画；真实报告中的数值、比例、标签和节点含义必须人工复核。')

# Shape[18]: 推荐工作流 → 更新
shape18_s2 = slide2.shapes[18]
shape18_s2.text = ('推荐工作流：万相负责高质感主视觉和技术仪表盘风格插图，即梦负责批量探索与中文说明图，'
                   'Z-Image 负责长文本和Logo场景，GLM 负责干净底图。')

# ============================================================
# Slide 3 - 数据集
# ============================================================
slide3 = prs.slides[2]

# Shape[13]: "98" → "112"
shape13_s3 = slide3.shapes[13]
shape13_s3.text = '112'

# Shape[14]: "含多个同模型候选图；五个模型均覆盖全部测试用例" → "六个测试模型"
shape14_s3 = slide3.shapes[14]
shape14_s3.text = '含多个同模型候选图；六个测试模型均覆盖全部测试用例（GPT 5.5 仅参照）'

# Shape[17]: "5" → "6+1（基准）"
shape17_s3 = slide3.shapes[17]
shape17_s3.text = '6+1'

# Shape[18]: "阿里、字节、百度、腾讯、智谱" → 加入阿里(Z-Image)
shape18_s3 = slide3.shapes[18]
shape18_s3.text = '阿里(万相+Z-Image)、字节、百度、腾讯、智谱'

# Shape[20]: TABLE 6rows x 3cols → add Z-Image row
table_s3 = slide3.shapes[20].table

# Add new row before last row (before GLM row)
new_row = table_s3.add_row()
# Copy GLM row formatting to new row (Z-Image)
new_row.cells[0].text = 'Z-Image'
new_row.cells[1].text = '18'
new_row.cells[2].text = '覆盖完整，长文本及Logo场景最优'

# Need to reorder: move new_row (last) to before GLM row
# Actually the table already had GLM as last, now we added one more. Let's swap.
# The table now has 7 rows: header, 万相, 即梦, 文心, 混元, GLM, new_row(Z-Image)
# We want: header, 万相, 即梦, 文心, 混元, GLM, Z-Image
# This is actually fine as-is since the new row is at the end

# But from the page 8-14, the order is: 万相, 即梦, 文心, 混元, GLM, Z-Image
# Let's also adjust the order in the table
# Current: 万相, 即梦, 文心, 混元, GLM, Z-Image (new) - this matches! Good.

# ============================================================
# Slide 4 - 成本对比
# ============================================================
slide4 = prs.slides[3]

table_s4 = slide4.shapes[3].table

# Add Z-Image row
new_row4 = table_s4.add_row()
new_row4.cells[0].text = 'Z-Image'
new_row4.cells[1].text = '约 0.05-0.08 元/张'
new_row4.cells[2].text = '较低'

# Update conclusion
shape5_s4 = slide4.shapes[5]
shape5_s4.text = ('结论：仅按本地资料换算，即梦成本最低，万相、Z-Image 次之；文心缺少图片额度，'
                  '混元成本明显偏高。')

# ============================================================
# Slide 5 - 综合评分
# ============================================================
slide5 = prs.slides[4]

# Shape[5]: 评分解释 → 更新
shape5_s5 = slide5.shapes[5]
shape5_s5.text = ('• 万相：新增 KPI/日志样本后综合领先，但文字、英文小字和伪指标仍需复核。\n'
                  '• 即梦：成本与中文可读性突出，适合批量方案探索和说明图。\n'
                  '• Z-Image：长文本、Logo 与数据图表场景表现最佳，视觉品质较高。\n'
                  '• GLM：中文与低干扰素材表现稳定，但视觉上限偏低。\n'
                  '• 文心、混元：本次样本中提示词贴合和文字控制不足。')

# Add Z-Image to the visual score cards
# Find last score card position and add new one
# The slide has picture boxes (Shape[7]) and text boxes
# Let's find the y of the last relevant text element
last_txt = slide5.shapes[5]
last_sp = last_txt._element
last_xfrm = last_sp.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
last_y = int(last_xfrm.find('{http://schemas.openxmlformats.org/drawingml/2006/main}off').get('y'))

# ============================================================
# Slide 6 - 场景优选矩阵
# ============================================================
slide6 = prs.slides[5]

# Shape[5]: 选择逻辑 → 更新
shape5_s6 = slide6.shapes[5]
shape5_s6.text = ('• 高质感成片：优先万相。\n'
                  '• 批量/中文说明图：优先即梦。\n'
                  '• 长文本/Logo/数据图表：优先 Z-Image。\n'
                  '• PPT 背景与干净底图：优先 GLM。\n'
                  '• 真实文字/数据/架构节点：必须人工校验或用结构化工具重绘。')

# ============================================================
# Slide 7 - 优势与风险
# ============================================================
slide7 = prs.slides[6]

table_s7 = slide7.shapes[3].table

# Add Z-Image row
new_row7 = table_s7.add_row()
new_row7.cells[0].text = 'Z-Image'
new_row7.cells[1].text = '长文本、Logo 与表格场景表现最佳，视觉品质较高'
new_row7.cells[2].text = '风格多样性有限，部分场景信息密度偏低'

# Save
output_file = '国产图片生成模型测试分析报告_已修改.pptx'
prs.save(output_file)
print(f'Saved to {output_file}')
print('All slides 1-7 updated successfully!')
