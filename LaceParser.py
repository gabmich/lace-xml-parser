import os
import xml.dom.minidom
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from .helpers import get_node_value


NEEDED_LACE_FILES = ['ImageReport_0.xml',]


class LaceParser():
    folder_path             = None
    xml_document            = None
    xml_items               = []
    evidences_ids           = set()
    evidences               = dict()
    template_for_gallery    = "./templates/basic_template_2.docx"
    template_for_list       = "./templates/basic_template_3.docx"

    def __init__(self, folder_path):
        self.check_folder(folder_path)
        self.parse_xml_document()
        self.get_all_items()
        self.extract_evidences()
        self.extract_items_content()

    def check_folder(self, folder_path):
        # chef if folder_path exists and ensure that it's a valid directory
        if os.path.exists(str(folder_path)) and os.path.isdir(str(folder_path)):
            print(folder_path, 'is a valid folder')
            self.check_presence_of_lace_files(folder_path)
            # self.get_folder_content()
            # self.list_and_rename_files()
        else:
            print('not a valid folder')

    def check_presence_of_lace_files(self, folder_path):
        files_list = os.listdir(folder_path)

        print('Files list of selected folder : ', files_list)
        if all(f in files_list for f in NEEDED_LACE_FILES):
            print('ok, all needed files are in the selected folder')
            self.folder_path = str(folder_path)
            return
        missing_files = [f for f in NEEDED_LACE_FILES if not f in files_list]
        print('Following files are missing in the selected folder : ', missing_files)


    def parse_xml_document(self):
        self.xml_document = xml.dom.minidom.parse(f'{self.folder_path}/ImageReport_0.xml')
        print(self.xml_document)

    def get_all_items(self):
        self.xml_items = self.xml_document.getElementsByTagName("Item")
        print(self.xml_items)

    def extract_evidences(self):
        for item in self.xml_items:
            evidence_id = item.getElementsByTagName("EvidenceID")[0].firstChild.nodeValue
            self.evidences_ids.add(evidence_id)
        print('EVIDENCES :', self.evidences_ids)

    def extract_items_content(self):
        for evidence in self.evidences_ids:
            items = []
            for item in self.xml_items:
                evidence_item = EvidenceItem(
                                    evidence_id=get_node_value(item,"EvidenceID"),
                                    file_id=get_node_value(item,"FileID"),
                                    image_path=get_node_value(item,"Thumbnail"),
                                    md5=get_node_value(item,"MD5"),
                                    partition=get_node_value(item,"Partition"),
                                    full_path=get_node_value(item,"FullPath"),
                                    file_name=get_node_value(item,"Filename"),
                                    created_at=get_node_value(item,"CreateDate"))
                items.append(evidence_item)

            # ICI : CONVERTIR CREATED_AT EN FORMAT DE DATE PLUS LISIBLE AU BESOIN
            print('ITEMS :', items)
            # export = DocxExporter(evidence_items=items, evidence=evidence)
            # print('EXPORT :', export)


class EvidenceItem():
    evidence_id             = ""
    file_id                 = ""
    image_path              = ""
    md5                     = ""
    partition               = ""
    full_path               = ""
    file_name               = ""
    created_at              = ""
    image                   = None

    def __init__(self,evidence_id,file_id,image_path,md5,partition,full_path,file_name,created_at):
        self.evidence_id = evidence_id
        self.file_id = file_id
        self.image_path = image_path
        self.md5 = md5
        self.partition = partition
        self.full_path = full_path
        self.file_name = file_name
        self.created_at = created_at

    @property
    def serialize(self):
        return {
            "file_id":self.file_id,
            "image_path":self.image_path,
            "md5":self.md5,
            "partition":self.partition,
            "full_path":self.full_path,
            "file_name":self.file_name,
            "image":self.image,
            "created_at":self.created_at,
        }


class DocxExporter():
    evidence_items          = []
    gallery_template_path   = "./templates/template_for_gallery.docx"
    gallery_docx            = None
    listing_template_path   = "./templates/basic_list_template.docx"
    listing_docx            = None
    start_date              = "Date de début"
    end_date                = "Date de fin"

    def __init__(self, evidence_items, evidence):
        self.sort_items_by_creation_date(evidence_items)
        self.create_docx_instances()
        context = { 'items' : self.evidence_items,
                    'evidence': evidence,
                    'start_date': self.start_date,
                    'end_date': self.end_date,}

        self.listing_docx.render(context)
        self.listing_docx.save(f"Evidence {evidence} list.docx")

        # self.gallery_docx.render(context)
        # self.gallery_docx.save(f"Evidence {evidence} images.docx")

    def create_docx_instances(self):
        self.listing_docx = DocxTemplate(self.listing_template_path)
        self.gallery_docx = DocxTemplate(self.gallery_template_path)

    def sort_items_by_creation_date(self, evidence_items):
        # self.evidence_items = sorted(evidence_items, key=lambda item:(item.created_at, item.created_at is None, item.created_at == ""))
        print('SORTED ITEMS :', self.evidence_items)

    def create_docx_img(self):
        image      = InlineImage("",
                                 image_descriptor=f"./lace_xml_export/Images/asdf",
                                 height=Mm(40))
