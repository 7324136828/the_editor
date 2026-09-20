from __future__ import annotations

import bisect
import re

from PySide6.QtGui import QTextCursor


LEXICON = """
ac-com-mo-da-tion ac-cu-rate ad-di-tion ad-min-is-tra-tion ad-van-tage
af-ter-noon al-go-rithm al-ign-ment al-ter-na-tive anal-y-sis an-nounce-ment
ap-pli-ca-tion ap-proach ar-chi-tec-ture ar-gu-ment ar-range-ment at-ten-tion
au-to-mat-ic back-ground bal-ance beau-ti-ful book-mark bound-ary
cal-cu-la-tion cam-era cap-a-bil-i-ty cat-e-go-ry cel-e-bra-tion char-ac-ter
col-lab-o-ra-tion col-lec-tion col-umn com-bi-na-tion com-mu-ni-ca-tion
com-par-i-son com-pat-i-bil-i-ty com-plex con-di-tion con-fig-u-ra-tion
con-nec-tion con-sid-er-a-tion con-tain-er con-tent con-tin-u-ous
con-trib-u-tion con-ver-sa-tion co-op-er-a-tion cor-rec-tion cre-ative
da-ta-base dec-o-ra-tion def-i-ni-tion de-pend-en-cy de-scrip-tion
de-sign de-vel-op-ment dic-tion-ar-y dif-fer-ence di-men-sion di-rec-tion
dis-trib-u-tion doc-u-ment doc-u-men-ta-tion dy-nam-ic ed-it-ing ed-u-ca-tion
ef-fi-cien-cy el-e-ment em-bed-ding en-gi-neer-ing en-vi-ron-ment equa-tion
es-sen-tial eval-u-a-tion ex-am-ple ex-cep-tion ex-pe-ri-ence ex-per-i-ment
ex-pla-na-tion ex-ten-sion fam-i-ly fea-ture fig-ure flex-i-ble float-ing
for-mat-ting func-tion gen-er-a-tion ge-om-e-try gram-mar graph-ic
head-ing hi-er-ar-chy his-to-ry hor-i-zon-tal hy-phen-ation iden-ti-ty
il-lus-tra-tion im-age im-ple-men-ta-tion im-por-tant in-cre-men-tal
in-de-pend-ent in-dex in-for-ma-tion in-ser-tion in-spec-tion in-stal-la-tion
in-struc-tion in-teg-ra-tion in-ter-ac-tion in-ter-face in-ter-na-tion-al
in-ter-val in-tro-duc-tion iso-la-tion jus-ti-fi-ca-tion knowl-edge
lan-guage lay-out li-brar-y lo-cal ma-chine man-age-ment mar-gin
math-e-mat-ics mea-sure-ment mem-o-ry meta-da-ta mi-gra-tion mod-i-fi-ca-tion
mul-ti-ple nav-i-ga-tion nec-es-sary num-ber ob-ject op-er-a-tion
op-por-tu-ni-ty op-ti-mi-za-tion or-gan-i-za-tion orig-i-nal out-line
pag-i-na-tion par-a-graph pa-ram-e-ter per-for-mance per-mis-sion
per-sis-tent pho-tog-ra-phy po-si-tion pre-de-fined prep-a-ra-tion
pre-sen-ta-tion pre-view pro-cess pro-duc-tion pro-gram pro-ject
pro-por-tion pro-tec-tion pub-li-ca-tion qual-i-ty ques-tion read-ing
rec-og-ni-tion rec-om-men-da-tion re-cov-ery ref-er-ence reg-is-tra-tion
re-la-tion-ship re-li-a-bil-i-ty ren-der-ing re-place-ment rep-re-sen-ta-tion
re-quire-ment res-o-lu-tion re-source re-sponse re-view re-vi-sion
se-cu-ri-ty se-lec-tion sem-an-tic sen-tence sep-a-ra-tion ser-vice
sig-na-ture sim-pli-fi-ca-tion soft-ware so-lu-tion spac-ing spec-i-fi-ca-tion
sta-tis-tics struc-ture sub-script sug-ges-tion su-per-script sup-port
sys-tem ta-ble tech-nol-o-gy tem-plate test-ing to-geth-er to-mor-row
tran-scrip-tion trans-for-ma-tion trans-par-ent ty-pog-ra-phy un-der-line
uni-code uni-ver-si-ty up-date us-a-bil-i-ty val-i-da-tion val-ue
vari-a-tion ver-sion ver-ti-cal vis-i-bil-i-ty wa-ter-mark win-dow
""".split()


def dictionary(entries=LEXICON):
    result = {}
    for entry in entries:
        parts = entry.split("-")
        word = "".join(parts)
        points, position = [], 0
        for part in parts[:-1]:
            position += len(part)
            if 2 <= position <= len(word) - 2:
                points.append(position)
        result[word] = tuple(points)
    return result


BREAKS = dictionary()


def apply_dictionary(document, hyphenate_caps=False, entries=None):
    positions = []
    lexicon = BREAKS if entries is None else dictionary(entries)
    text = document.toPlainText()
    for match in re.finditer(r"\b[A-Za-z]{5,}\b", text):
        word = match.group()
        if word.isupper() and not hyphenate_caps:
            continue
        start = len(text[:match.start()].encode("utf-16-le")) // 2
        positions.extend(start + point for point in lexicon.get(word.lower(), ()))
    edit = QTextCursor(document)
    edit.beginEditBlock()
    for position in reversed(positions):
        edit.setPosition(position)
        edit.insertText("\u00ad")
    edit.endEditBlock()
    return positions


def layout_position(position, inserted):
    return position + bisect.bisect_right(inserted, position)
