# -*- coding: utf-8 -*-
"""깨진 워드(.docx) 되살리기.

워드 파일(.docx)은 여러 조각(본문·서식·그림 …)을 묶은 압축 파일이고, 맨 끝에
'어느 조각이 어디 있는지' 적힌 목차가 붙는다. 메일·다운로드 중에 **파일 끝이 잘리면**
그 목차가 없어져 프로그램이 파일을 못 연다(실제로 겪음: 부산 남포동 워드 원본).

여기서는 파일 안에 남아 있는 조각을 처음부터 순서대로 주워 다시 묶는다.
잘린 뒤쪽에 있던 조각은 **파일 안에 아예 없으므로** 되살릴 수 없다 — 그 목록을 돌려줘서
화면에 알린다. 없는 조각을 가리키는 연결은 지워야 워드가 연다.

기존 IM 도구의 preview._salvage_docx_zip 을 바탕으로, '없는 조각 연결 정리' 를 더했다.
(그것만으로는 부산 남포동이 docProps/app.xml 이 없다며 여전히 안 열렸다 — 2026-10-07 확인)
"""
import io
import re
import struct
import zipfile
import zlib
import posixpath

# 이 조각이 빠지면 무엇이 달라지는지 — 화면 안내용
_PART_MEANING = {
    "word/styles.xml": "서식(글꼴·글씨 크기·제목 모양)",
    "word/numbering.xml": "번호·글머리표 매기기",
    "word/theme/theme1.xml": "기본 색·기본 글꼴(테마)",
    "word/settings.xml": "문서 설정",
    "word/fontTable.xml": "글꼴 목록",
    "word/webSettings.xml": "웹 설정",
    "docProps/app.xml": "문서 정보(쪽수 등)",
    "docProps/core.xml": "문서 정보(작성자·날짜)",
}


def meaning_of(part):
    if part.startswith("word/media/"):
        return "그림"
    if part.startswith("word/header"):
        return "머리말"
    if part.startswith("word/footer"):
        return "꼬리말"
    if part.startswith("customXml/"):
        return "부가 정보(보이지 않음)"
    return _PART_MEANING.get(part, "기타")


def _local_entries(data):
    """파일 앞에서부터 조각 머리(PK 03 04)를 따라가며 온전한 조각만 꺼낸다."""
    entries, cut = [], []
    i = 0
    while True:
        j = data.find(b"PK\x03\x04", i)
        if j < 0 or j + 30 > len(data):
            break
        try:
            (_ver, flag, method, _mt, _md, _crc,
             csize, usize, fnl, efl) = struct.unpack("<HHHHHIIIHH", data[j + 4:j + 30])
        except struct.error:
            break
        name = data[j + 30:j + 30 + fnl].decode("utf-8", "replace")
        body = j + 30 + fnl + efl
        if (flag & 0x08) or (csize == 0 and usize == 0 and method != 0):
            cut.append(name)                 # 크기가 머리에 없는 형식 → 여기서 더 못 읽는다
            break
        comp = data[body:body + csize]
        if len(comp) < csize:                # 파일이 이 조각 중간에서 잘림
            cut.append(name)
            break
        try:
            entries.append((name, comp if method == 0 else zlib.decompress(comp, -15)))
        except Exception:
            cut.append(name)                 # 이 조각만 깨짐 — 버리고 계속
        i = body + csize
    return entries, cut


def _resolve(rels_name, target):
    """rels 파일 기준 상대 경로 → 패키지 안 경로."""
    if target.startswith("/"):
        return target.lstrip("/")
    owner_dir = posixpath.dirname(posixpath.dirname(rels_name))      # 'word/_rels/x.rels' → 'word'
    return posixpath.normpath(posixpath.join(owner_dir, target)).lstrip("./")


_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _drop_dangling_pictures(parts):
    """연결이 사라진 rId 를 쓰는 그림(w:drawing / w:pict)을 통째로 뺀다."""
    from lxml import etree
    for name in list(parts):
        if not (name.startswith("word/") and name.endswith(".xml")) or "/_rels/" in name:
            continue
        rels = posixpath.join(posixpath.dirname(name), "_rels", posixpath.basename(name) + ".rels")
        if rels not in parts:
            continue
        ids = set(re.findall(r'Id="([^"]+)"', parts[rels].decode("utf-8", "replace")))
        try:
            root = etree.fromstring(parts[name])
        except Exception:
            continue
        changed = False
        for el in list(root.iter()):
            bad = any(k.startswith("{%s}" % _R_NS) and v not in ids for k, v in el.attrib.items())
            if not bad:
                continue
            # 가장 가까운 그림 덩어리를 찾아 뺀다(없으면 그 속성만 뺀다)
            box = el
            while box is not None and etree.QName(box).localname not in ("drawing", "pict", "object"):
                box = box.getparent()
            if box is not None and box.getparent() is not None:
                box.getparent().remove(box)
            else:
                for k in [k for k in el.attrib if k.startswith("{%s}" % _R_NS)]:
                    del el.attrib[k]
            changed = True
        if changed:
            parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def repair_docx(data):
    """깨진 docx 를 되살린다.

    반환: (새 docx bytes 또는 None, 빠진 조각 목록[경로])
      · None 이면 본문(word/document.xml)조차 없어 살릴 수 없는 것.
    """
    entries, _cut = _local_entries(data)
    parts = {}
    for name, raw in entries:
        parts.setdefault(name, raw)          # 같은 이름이 또 나오면 앞의 것
    if "word/document.xml" not in parts:
        return None, []

    missing = set()
    # 없는 조각을 가리키는 연결 지우기
    for rels in [n for n in parts if n.endswith(".rels")]:
        xml = parts[rels].decode("utf-8", "replace")

        def keep(m):
            rel = m.group(0)
            if 'TargetMode="External"' in rel:
                return rel
            t = re.search(r'Target="([^"]+)"', rel)
            if not t:
                return rel
            p = _resolve(rels, t.group(1))
            if p in parts:
                return rel
            missing.add(p)
            return ""
        parts[rels] = re.sub(r"<Relationship\b[^>]*/>", keep, xml).encode("utf-8")

    # 본문·머리말·꼬리말이 지운 연결(그림 등)을 아직 가리키면 워드가 PDF 저장 등에서
    # '명령을 정상적으로 종료할 수 없습니다' 로 멈춘다(실제로 겪음) → 그 그림 자리를 비운다.
    _drop_dangling_pictures(parts)

    # 목록표([Content_Types].xml)에서도 없는 조각을 뺀다
    ct = parts.get("[Content_Types].xml")
    if ct is not None:
        xml = ct.decode("utf-8", "replace")
        xml = re.sub(r'<Override\b[^>]*PartName="/([^"]+)"[^>]*/>',
                     lambda m: m.group(0) if m.group(1) in parts else "", xml)
        parts["[Content_Types].xml"] = xml.encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, raw in parts.items():
            z.writestr(name, raw)
    return out.getvalue(), sorted(missing)


# 결과 모양에 영향이 없는 조각 — 화면 안내에서는 뺀다(괜히 겁주지 않게)
_HARMLESS = ("docProps/", "customXml/", "word/webSettings.xml", "word/settings.xml",
             "word/fontTable.xml")


def missing_summary(missing):
    """빠진 조각 목록을 사람 말로 — '그림 4장, 서식(글꼴·글씨 크기·제목 모양), …'"""
    pics = sum(1 for p in missing if p.startswith("word/media/"))
    rest = []
    for p in missing:
        if p.startswith("word/media/") or p.startswith(_HARMLESS):
            continue
        m = meaning_of(p)
        if m not in rest:
            rest.append(m)
    items = ([f"그림 {pics}장"] if pics else []) + rest
    return ", ".join(items)
