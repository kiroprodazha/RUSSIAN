#!/usr/bin/env python3
"""
Generate a valid .pptx presentation file for Blok's poem analysis.
No external dependencies — builds PPTX (Office Open XML) from scratch.
"""

import zipfile
import os

OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "blok_presentation.pptx")

# ─── Slide content ───────────────────────────────────────────────────────────

slides_data = [
    {
        "title": "«Предчувствую Тебя.\nГода проходят мимо…»",
        "body": "А. А. Блок\nЦикл «Стихи о Прекрасной Даме» (1901–1902)\n\nАнализ стихотворения",
    },
    {
        "title": "Текст стихотворения",
        "body": (
            "Предчувствую Тебя. Года проходят мимо —\n"
            "Всё в облике одном предчувствую Тебя.\n"
            "Весь горизонт в огне — и ясен нестерпимо,\n"
            "И молча жду, — тоскуя и любя.\n\n"
            "Весь горизонт в огне, и близко появленье,\n"
            "Но страшно мне: изменишь облик Ты,\n"
            "И дерзкое возбудишь подозренье,\n"
            "Сменив в конце привычные черты.\n\n"
            "О, как паду — и горестно, и низко,\n"
            "Не одолев смертельныя мечты!\n"
            "Как ясен горизонт! И лучезарность близко.\n"
            "Но страшно мне: изменишь облик Ты.\n\n"
            "1901"
        ),
    },
    {
        "title": "Идея произведения",
        "body": (
            "Стихотворение — поэтический манифест раннего Блока, ключ ко всему циклу «Стихи о Прекрасной Даме».\n\n"
            "Главная идея: мистическое предчувствие встречи с Вечной Женственностью — божественным идеалом — "
            "соединяется со страхом её земного воплощения, в котором небесный облик может исказиться.\n\n"
            "Любовь у Блока — не земное чувство, а религиозное служение; лирический герой — рыцарь, монах, "
            "ждущий Откровения. Возлюбленная превращается в Идеал, в Святую."
        ),
    },
    {
        "title": "Аргумент 1. Образ Прекрасной Дамы — мистический символ",
        "body": (
            "• Местоимения «Тебя», «Ты» написаны с заглавной буквы — как обращение к Богу. "
            "Это не земная женщина, а воплощение Души Мира (идея Вл. Соловьёва).\n\n"
            "• Строка «Всё в облике одном предчувствую Тебя» — герой видит Её во всём; образ растворён в мире.\n\n"
            "• Приёмы: сакральная заглавная буква, повторы-заклинания, недоговорённость, символика.\n\n"
            "Цитата: «Предчувствую Тебя. Года проходят мимо — / Всё в облике одном предчувствую Тебя»."
        ),
    },
    {
        "title": "Аргумент 2. Двоемирие: ожидание и страх",
        "body": (
            "• Лирический герой одновременно жаждет встречи и боится её: "
            "«молча жду, — тоскуя и любя» — оксюморонное соединение чувств.\n\n"
            "• Страх рождён предчувствием искажения идеала: «страшно мне: изменишь облик Ты».\n\n"
            "• Приёмы: антитеза («ясен» — «страшно», «лучезарность» — «паду низко»), "
            "оксюморон, повтор-рефрен, нагнетающий тревогу.\n\n"
            "Цитата: «И молча жду, — тоскуя и любя… / Но страшно мне: изменишь облик Ты»."
        ),
    },
    {
        "title": "Аргумент 3. Космический пейзаж как образ Откровения",
        "body": (
            "• Пейзаж символический: «Весь горизонт в огне» — образ зари, апокалиптического света, "
            "грядущего явления Вечной Женственности.\n\n"
            "• «Ясен нестерпимо» — оксюморон: свет настолько силён, что человек не выдерживает Откровения.\n\n"
            "• Приёмы: цветопись (огонь, лучезарность), символ зари, анафора «Весь горизонт в огне», "
            "церковнославянизм «смертельныя» — молитвенная интонация.\n\n"
            "Цитата: «Весь горизонт в огне — и ясен нестерпимо… / Как ясен горизонт! И лучезарность близко»."
        ),
    },
    {
        "title": "Заключение",
        "body": (
            "Стихотворение «Предчувствую Тебя…» — квинтэссенция мировоззрения молодого Блока: "
            "мир — храм, любовь — служение, земная встреча — лишь отблеск вечного.\n\n"
            "Через символику зари, сакральные местоимения, антитезу ожидания и страха поэт передаёт "
            "напряжённое предчувствие чуда — и трагическую неизбежность его земного искажения.\n\n"
            "Идеи Блока живы и сегодня: каждый, кто по-настоящему любит, превращает любимого человека "
            "в свой идеал, ждёт его как чуда — и боится потерять."
        ),
    },
    {
        "title": "Стихотворение собственного сочинения",
        "body": (
            "В традициях цикла «Стихи о Прекрасной Даме»\n\n"
            "[Здесь будет ваше стихотворение,\n"
            "посвящённое возлюбленной,\n"
            "написанное в духе А. А. Блока]\n\n"
            "(8–12 строк)"
        ),
    },
]

# ─── PPTX XML templates ─────────────────────────────────────────────────────

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  {slide_overrides}
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"""

PRESENTATION_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    {slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""

PRESENTATION_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  {slide_rels}
  <Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""

SLIDE_MASTER = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="1A1530"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
    accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>"""

SLIDE_MASTER_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""

SLIDE_LAYOUT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr/>
  </p:spTree></p:cSld>
</p:sldLayout>"""

SLIDE_LAYOUT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""

THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="BlokTheme">
  <a:themeElements>
    <a:clrScheme name="Blok">
      <a:dk1><a:srgbClr val="1A1530"/></a:dk1>
      <a:lt1><a:srgbClr val="F5F1E8"/></a:lt1>
      <a:dk2><a:srgbClr val="2A1F4A"/></a:dk2>
      <a:lt2><a:srgbClr val="E0D4F5"/></a:lt2>
      <a:accent1><a:srgbClr val="F0D27A"/></a:accent1>
      <a:accent2><a:srgbClr val="D9A6C2"/></a:accent2>
      <a:accent3><a:srgbClr val="FFE8B0"/></a:accent3>
      <a:accent4><a:srgbClr val="B8ACD1"/></a:accent4>
      <a:accent5><a:srgbClr val="3A1F2E"/></a:accent5>
      <a:accent6><a:srgbClr val="F0D27A"/></a:accent6>
      <a:hlink><a:srgbClr val="F0D27A"/></a:hlink>
      <a:folHlink><a:srgbClr val="D9A6C2"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Blok">
      <a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>"""


def make_slide_xml(title: str, body: str) -> str:
    """Create slide XML with title and body text boxes."""
    # Escape XML special chars
    def esc(s):
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

    title_esc = esc(title)
    
    # Build body paragraphs
    body_paras = ""
    for line in body.split("\n"):
        line_esc = esc(line)
        body_paras += f'<a:p><a:r><a:rPr lang="ru-RU" sz="1600" dirty="0"/><a:t>{line_esc}</a:t></a:r></a:p>\n'

    # Build title paragraphs (support multiline titles)
    title_paras = ""
    for line in title.split("\n"):
        line_esc = esc(line)
        title_paras += f'<a:p><a:r><a:rPr lang="ru-RU" sz="2800" b="1" i="1" dirty="0"><a:solidFill><a:srgbClr val="F0D27A"/></a:solidFill></a:rPr><a:t>{line_esc}</a:t></a:r></a:p>\n'

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="1A1530"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      <!-- Title -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="600000" y="300000"/><a:ext cx="10900000" cy="1200000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          {title_paras}
        </p:txBody>
      </p:sp>
      <!-- Body -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="600000" y="1600000"/><a:ext cx="10900000" cy="4800000"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="t"/>
          <a:lstStyle/>
          {body_paras}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>"""


def make_slide_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def build_pptx(output_path: str):
    n = len(slides_data)

    # Build dynamic parts
    slide_overrides = "\n".join(
        f'  <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(n)
    )
    slide_ids = "\n".join(
        f'    <p:sldId id="{256+i}" r:id="rId{10+i}"/>'
        for i in range(n)
    )
    slide_pres_rels = "\n".join(
        f'  <Relationship Id="rId{10+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>'
        for i in range(n)
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES.format(slide_overrides=slide_overrides))
        zf.writestr("_rels/.rels", RELS_ROOT)
        zf.writestr("ppt/presentation.xml", PRESENTATION_XML.format(slide_ids=slide_ids))
        zf.writestr("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS.format(slide_rels=slide_pres_rels))
        zf.writestr("ppt/slideMasters/slideMaster1.xml", SLIDE_MASTER)
        zf.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", SLIDE_MASTER_RELS)
        zf.writestr("ppt/slideLayouts/slideLayout1.xml", SLIDE_LAYOUT)
        zf.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", SLIDE_LAYOUT_RELS)
        zf.writestr("ppt/theme/theme1.xml", THEME)

        for i, slide in enumerate(slides_data):
            slide_xml = make_slide_xml(slide["title"], slide["body"])
            zf.writestr(f"ppt/slides/slide{i+1}.xml", slide_xml)
            zf.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", make_slide_rels())

    print(f"Created: {output_path}")
    print(f"Slides: {n}")


if __name__ == "__main__":
    build_pptx(OUTPUT)
