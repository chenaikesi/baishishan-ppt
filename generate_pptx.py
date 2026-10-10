from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pathlib import Path
import requests
import os

OUTPUT = Path('baishishan_presentation.pptx')
ASSET_DIR = Path('assets')
ASSET_DIR.mkdir(exist_ok=True)

slides_data = [
    {
        'title': '白石山',
        'subtitle': '自然·奇观·徒步圣地',
        'image': 'https://images.unsplash.com/photo-1501785888041-af3ef285b470?auto=format&fit=crop&w=1400&q=80',
        'notes': '白石山以奇石、云海、古松和山景闻名，是适合徒步和摄影的自然景区。'
    },
    {
        'title': '目录',
        'bullets': ['白石山简介', '地理与位置', '地质与地貌', '主要景点', '植被与野生动物', '徒步与活动推荐', '交通与住宿', '旅行贴士'],
        'image': 'https://images.unsplash.com/photo-1441974231531-c6227db76b6e?auto=format&fit=crop&w=1400&q=80',
        'notes': '本页概括本次演示的结构和重点内容。'
    },
    {
        'title': '白石山简介',
        'bullets': ['位于河北省境内，属于典型的山地风景区', '以白色岩石、奇峰和林木景观著称', '适合登山、摄影和休闲郊游'],
        'image': 'https://images.unsplash.com/photo-1464822759023-fed622ff2c3b?auto=format&fit=crop&w=1400&q=80',
        'notes': '强调白石山的吸引力：自然景观和适合休闲探险的特点。'
    },
    {
        'title': '地理位置',
        'bullets': ['交通便捷，适合周末和假期出行', '具备明显的山地地形和季节变化', '春夏秋冬都可以欣赏不同风貌'],
        'image': 'https://images.unsplash.com/photo-1527631746610-bca00a040d60?auto=format&fit=crop&w=1400&q=80',
        'notes': '从地理位置和可达性介绍景区的出游便利性。'
    },
    {
        'title': '地质与地貌',
        'bullets': ['以花岗岩地貌为主，长期风化形成独特石面', '岩层、裂隙和峰脊形成独特景观', '地质景观与自然风光兼具'],
        'image': 'https://images.unsplash.com/photo-1501854140801-50d01698950b?auto=format&fit=crop&w=1400&q=80',
        'notes': '在不使用过多科普术语的前提下，解释地形成因。'
    },
    {
        'title': '主要景点（一）',
        'bullets': ['观景台：适合眺望全景和日出', '奇石区：形态各异的石峰和岩层', '古松：代表自然生态与历史感'],
        'image': 'https://images.unsplash.com/photo-1500530855697-b586d89ba3ee?auto=format&fit=crop&w=1400&q=80',
        'notes': '突出视觉冲击和摄影价值。'
    },
    {
        'title': '主要景点（二）',
        'bullets': ['栈道与步道：适合徒步体验', '溪流与瀑布：增添山野气息', '观景点与摄影点：适合记录旅行风景'],
        'image': 'https://images.unsplash.com/photo-1472396961693-142e6e269027?auto=format&fit=crop&w=1400&q=80',
        'notes': '补充体验式景点，让游客根据兴趣选择路线。'
    },
    {
        'title': '植被与野生动物',
        'bullets': ['以山地森林和灌丛为主', '春夏季野花丰富，秋季色彩艳丽', '适合观察鸟类和小型野生动物'],
        'image': 'https://images.unsplash.com/photo-1506744038136-46273834b3fb?auto=format&fit=crop&w=1400&q=80',
        'notes': '从生态维度强调白石山除了山景还有丰富自然环境。'
    },
    {
        'title': '徒步与活动推荐',
        'bullets': ['初级路线：适合家庭和休闲游客', '中级路线：适合喜欢徒步的游客', '摄影路线：适合日出、云海和山景拍摄'],
        'image': 'https://images.unsplash.com/photo-1524661135-423995f22d0b?auto=format&fit=crop&w=1400&q=80',
        'notes': '建议游客根据体能和兴趣选择不同路线。'
    },
    {
        'title': '交通与住宿',
        'bullets': ['自驾或公共交通均可到达', '周边有民宿、酒店和服务设施', '提前预订更稳妥，节假日需注意车位与客流'],
        'image': 'https://images.unsplash.com/photo-1501785888041-af3ef285b470?auto=format&fit=crop&w=1400&q=80',
        'notes': '让游客对行程规划和住宿安排有更清晰认知。'
    },
    {
        'title': '旅行贴士',
        'bullets': ['最佳季节：春季和秋季较舒适', '注意保暖、防滑和水分补充', '遵守景区规定，保护环境和野生动物'],
        'image': 'https://images.unsplash.com/photo-1501555088652-021faa106b9b?auto=format&fit=crop&w=1400&q=80',
        'notes': '结尾以实用建议和环保提醒收束。'
    },
    {
        'title': '结尾',
        'bullets': ['白石山是一处融合自然、地貌和徒步体验的山地景区', '适合放松、拍照、探险和短途旅行', '旅行的意义在于感受自然与慢下来'],
        'image': 'https://images.unsplash.com/photo-1501785888041-af3ef285b470?auto=format&fit=crop&w=1400&q=80',
        'notes': '以优雅的尾句收束，适合最后一页留白和问答。'
    },
]


def download_image(url: str, filename: str) -> str:
    dest = ASSET_DIR / filename
    if dest.exists():
        return str(dest)
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        dest.write_bytes(resp.content)
    except Exception:
        # 不让下载失败中断输出
        return ''
    return str(dest)


def add_background(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def set_title(slide, title, subtitle=None):
    title_box = slide.shapes.title
    if title_box is None:
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.35), Inches(11.5), Inches(0.8))
    tf = title_box.text_frame
    tf.text = title
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.name = 'Microsoft YaHei'
    tf.paragraphs[0].font.color.rgb = RGBColor(22, 38, 55)
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.0), Inches(11.0), Inches(0.5))
        tf2 = sub_box.text_frame
        p = tf2.paragraphs[0]
        p.text = subtitle
        p.alignment = PP_ALIGN.LEFT
        p.font.size = Pt(14)
        p.font.name = 'Microsoft YaHei'
        p.font.color.rgb = RGBColor(88, 103, 117)


def add_bullets(slide, bullets):
    box = slide.shapes.add_textbox(Inches(0.7), Inches(1.5), Inches(5.2), Inches(4.8))
    tf = box.text_frame
    tf.word_wrap = True
    for idx, item in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.level = 0
        p.font.size = Pt(22)
        p.font.name = 'Microsoft YaHei'
        p.font.bold = False
        p.bullet = True
        p.space_after = Pt(10)
        p.alignment = PP_ALIGN.LEFT
        p.font.color.rgb = RGBColor(32, 46, 58)


def add_picture(slide, image_file, left, top, width, height):
    if not image_file:
        return
    try:
        slide.shapes.add_picture(image_file, left, top, width, height)
    except Exception:
        pass


def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for idx, item in enumerate(slides_data):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(slide, RGBColor(245, 248, 250))

        if idx == 0:
            # 封面，大图背景 + 半透明覆盖层
            image_file = download_image(item['image'], f'cover_{idx}.jpg')
            if image_file:
                slide.shapes.add_picture(image_file, 0, 0, prs.slide_width, prs.slide_height)
            overlay = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
            fill = overlay.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(15, 22, 31)
            fill.transparency = 0.35
            overlay.line.fill.background()
            title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.5), Inches(1.3))
            tf = title_box.text_frame
            p = tf.paragraphs[0]
            p.text = item['title']
            p.font.size = Pt(32)
            p.font.bold = True
            p.font.name = 'Microsoft YaHei'
            p.font.color.rgb = RGBColor(255, 255, 255)
            sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), Inches(7.0), Inches(0.6))
            tf2 = sub_box.text_frame
            p2 = tf2.paragraphs[0]
            p2.text = item['subtitle']
            p2.font.size = Pt(18)
            p2.font.name = 'Microsoft YaHei'
            p2.font.color.rgb = RGBColor(220, 232, 245)
            continue

        # 普通页：文本左，图片右
        if 'bullets' in item:
            add_bullets(slide, item['bullets'])
        else:
            # 无列表页（如封面已处理）
            pass

        # 标题
        title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.35), Inches(6.2), Inches(0.8))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = item['title']
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.name = 'Microsoft YaHei'
        p.font.color.rgb = RGBColor(20, 32, 42)

        image_file = download_image(item['image'], f'slide_{idx}.jpg')
        if image_file:
            add_picture(slide, image_file, Inches(7.0), Inches(1.5), Inches(5.3), Inches(4.8))

    prs.save(OUTPUT)
    print(f'PPT 已生成：{OUTPUT.resolve()}')


if __name__ == '__main__':
    build_presentation()
