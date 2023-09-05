from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import CT_Inline, parse_xml, CT_Picture
from docx.shape import InlineShape
from docx.shared import Length
from docx.text.run import Run


# noinspection PyProtectedMember
def add_linked_pic(r: Run, image_path: str, width: Length, height: Length) -> InlineShape:
    """
    Image will be inserted as character in the Run.
    :param r: current run
    :param image_path:
        Seems like it has to be absolute path.as_uri() like "file:///full/path/file.jpg".
        It also works with relative path like "./folder/image.jpg"
    :param width: size of image in document
    :param height: size of image in document
    """

    # create RELATION
    relations = r.part.rels
    rel_id = relations._next_rId
    relations.add_relationship(reltype=RELATIONSHIP_TYPE.IMAGE, target=image_path, rId=rel_id, is_external=True)

    # Comment about pic_id from python-docx creators:
    # -- Word doesn't seem to use this, but does not omit it
    pic_id = 0

    # Next code taken from this method:
    # def new(cls, pic_id, filename, rId, cx, cy):
    # Just one line changed in order to replace `r:embed` with `r:link`.

    # The following fout lines created to make variable names same as in python-docx method.
    filename = image_path  # Filename - something useless. will make it equal to image_path
    cx = width
    cy = height

    # Expand that code as CT_Picture.new(pic_id, filename, rId, cx, cy):
    pic = parse_xml(CT_Picture._pic_xml())
    pic.nvPicPr.cNvPr.id = pic_id
    pic.nvPicPr.cNvPr.name = filename

    # pic.blipFill.blip.embed = rId  # This line is replaced with next one
    pic.blipFill.blip.link = rel_id

    pic.spPr.cx = cx
    pic.spPr.cy = cy

    shape_id = r.part.next_id

    # Now from here: inline = cls.new(cx, cy, shape_id, pic)
    inline = CT_Inline.new(cx, cy, shape_id, pic)
    inline = r._r.add_drawing(inline)

    return InlineShape(inline)
