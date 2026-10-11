"""Owned regressions for External metadata referencing package members."""
from io import BytesIO
import pytest
from aspose.words_foss.dom import DocxDocument
from aspose.words_foss.dom.comment_utc import ROLES,elements,REL
from .test_dom_comment_metadata_relationships import conflicting
from .test_dom_comment_metadata import comment

@pytest.mark.parametrize('role', list(ROLES))
@pytest.mark.parametrize('flat', [False, True])
def test_relative_external_metadata_preserves_selected_date(tmp_path, role, flat):
    doc=conflicting(tmp_path,role,'rId0','rId9')
    root=doc._package.tree('word/_rels/document.xml.rels').documentElement
    link=next(n for n in elements(root,REL,'Relationship') if n.getAttribute('Type')==ROLES[role][2] and n.getAttribute('Id')=='rId0')
    link.setAttribute('TargetMode','External')
    doc._package._dirty.add('word/_rels/document.xml.rels')
    doc=DocxDocument(BytesIO(doc._package.to_bytes()))
    held=comment(doc)
    # Native 26.9 uses the available package member, not the legacy 2026 date
    # or later typed relation's 2044 date. No external access is involved.
    assert held.date_time_utc.year==2030
    before={name:doc._package.payload(name) for name in doc.part_names}
    data=doc.to_flat_opc() if flat else doc.to_bytes()
    assert comment(DocxDocument(BytesIO(data))).date_time_utc.year==2030
    assert {name:doc._package.payload(name) for name in doc.part_names}==before
    out=tmp_path/'saved'
    for _ in range(2):
        (doc.save_flat_opc if flat else doc.save)(out)
        assert comment(doc)is held and held.date_time_utc.year==2030
        cold=DocxDocument(out)
        assert comment(cold).date_time_utc.year==2030
        relations=cold._package.tree('word/_rels/document.xml.rels').documentElement
        for _,(_,_,typ) in ROLES.items():
            assert sum(n.getAttribute('Type')==typ for n in elements(relations,REL,'Relationship'))==1


@pytest.mark.parametrize('target', ['https://example.invalid/a', '../../outside.xml',
                                  'file:///outside.xml', 'comments.xml#fragment'])
def test_external_invalid_target_does_not_resolve_outside_package(target):
    from xml.dom.minidom import Document
    from aspose.words_foss.dom.comment_utc import package_target
    link = Document().createElementNS(REL, 'Relationship')
    link.setAttribute('Target', target)
    link.setAttribute('TargetMode', 'External')
    assert package_target(link, 'word/document.xml') is None
    link.setAttribute('TargetMode', 'Internal')
    with pytest.raises(ValueError):
        package_target(link, 'word/document.xml')


def test_external_retained_edge_preserves_losing_metadata_member(tmp_path):
    doc = conflicting(tmp_path, 'commentsIds', 'rId0', 'rId9')
    rels = doc._package.tree('word/_rels/document.xml.rels').documentElement
    link = rels.ownerDocument.createElementNS(REL, 'Relationship')
    link.setAttribute('xmlns', REL)
    for name, value in [('Id', 'owned-keep'), ('Type', 'urn:owned:keep'),
                        ('Target', 'owned-second-commentsIds.xml'), ('TargetMode', 'External')]:
        link.setAttribute(name, value)
    rels.appendChild(link)
    doc._package._dirty.add('word/_rels/document.xml.rels')
    doc = DocxDocument(BytesIO(doc._package.to_bytes()))
    cold = DocxDocument(BytesIO(doc.to_bytes()))
    assert 'word/owned-second-commentsIds.xml' in cold.part_names
    assert comment(cold).date_time_utc.year == 2030
