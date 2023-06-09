import os
import xml.dom.minidom
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from resources.helpers import (get_node_value,
                               create_dt,
                               debug,
                               to_chunks,)


NEEDED_LACE_FILES = ["ImageReport_0.xml",]
LACE_DT_FORMAT    = "%a, %b %d, %Y %H:%M:%S %Z"
EXPORT_PATH       = "../output"


class LaceParser():
    folder_path             = None
    xml_document            = None
    xml_items               = []
    extracted_items         = []
    evidences_ids           = set()
    evidences               = dict()
    template_for_gallery    = "./templates/basic_template_2.docx"
    template_for_list       = "./templates/basic_template_3.docx"

    def __init__(self, folder_path):
        self.check_folder(folder_path)
        self.parse_xml_document()
        self.extract_all_items()
        self.extract_evidences()
        self.extract_items_content()
        self.export_to_templates()

    def check_folder(self, folder_path):
        # checks if folder_path exists and ensure that it's a valid directory
        if os.path.exists(str(folder_path)) and os.path.isdir(str(folder_path)):
            print(folder_path, "is a valid folder")
            self.check_presence_of_lace_files(folder_path)
            # self.get_folder_content()
            # self.list_and_rename_files()
        else:
            print("not a valid folder")

    def check_presence_of_lace_files(self, folder_path):
        # checks if all needed files are in the selected folder
        files_list = os.listdir(folder_path)

        print("Files list of selected folder :", files_list)
        if all(f in files_list for f in NEEDED_LACE_FILES):
            print("ok, all needed files are in the selected folder")
            self.folder_path = str(folder_path)
            return
        missing_files = [f for f in NEEDED_LACE_FILES if not f in files_list]
        print("Following files are missing in the selected folder :", missing_files)


    def parse_xml_document(self):
        self.xml_document = xml.dom.minidom.parse(f"{self.folder_path}/ImageReport_0.xml")
        print(self.xml_document)

    def extract_all_items(self):
        self.xml_items = self.xml_document.getElementsByTagName("Item")
        print("NOMBRE D'ITEMS TROUVES :", len(self.xml_items))

    def extract_evidences(self):
        for item in self.xml_items:
            evidence_id = get_node_value(item, "EvidenceID") # item.getElementsByTagName("EvidenceID")[0].firstChild.nodeValue
            self.evidences_ids.add(evidence_id)
        self.evidences = {evidence_id:[] for evidence_id in self.evidences_ids}
        print("EVIDENCES IDS :", self.evidences_ids)
        print("EVIDENCES DICT :", self.evidences)

    def extract_items_content(self):
        for evidence in self.evidences:
            for item in self.xml_items:
                evidence_id = get_node_value(item,"EvidenceID")
                if evidence_id == evidence:
                    created_at = create_dt(get_node_value(item,"CreateDate"), LACE_DT_FORMAT)
                    updated_at = create_dt(get_node_value(item,"ModifyDate"), LACE_DT_FORMAT)

                    evidence_item = {
                        "evidence_id": evidence_id,
                        "file_id": get_node_value(item,"FileID"),
                        "image_path": get_node_value(item,"Thumbnail"),
                        "image_full_path": f"{self.folder_path}/{get_node_value(item,'Thumbnail')}",
                        "md5": get_node_value(item,"MD5"),
                        "partition": get_node_value(item,"Partition"),
                        "full_path": get_node_value(item,"FullPath"),
                        "file_name": get_node_value(item,"Filename"),
                        "created_at": created_at,
                        "updated_at": updated_at,
                        "folder_path": self.folder_path,
                        "image": None,
                    }

                    self.evidences[evidence].append(evidence_item)

            # ICI : CONVERTIR CREATED_AT EN FORMAT DE DATE PLUS LISIBLE AU BESOIN
            print(f"NOMBRE D'ITEMS CREES POUR EVIDENCE '{evidence}' :", len(self.evidences[evidence]))
            # export = DocxExporter(evidence_items=items, evidence=evidence)
            # print('EXPORT :', export)

    def export_to_templates(self):
        for i, evidence in enumerate(self.evidences):
            export = DocxExporter(evidence_items=self.evidences[evidence], evidence=evidence, count=i+1)
            print(evidence, "EXPORTED !")


class DocxExporter():
    evidence                = ""
    evidence_items          = []
    count                   = 1

    gallery_template_path   = "./resources/templates/Images/template_for_gallery.docx"
    listing_template_path   = "./resources/templates/Images/template_for_list.docx"

    gallery_docx            = None
    listing_docx            = None

    context                 = dict()
    
    start_date              = "Date de début"
    end_date                = "Date de fin"

    def __init__(self, evidence_items, evidence, count):
        self.evidence = evidence
        self.evidence_items = evidence_items
        self.count = count
        
        self.sort_items_by_date()
        self.create_docx_instances()
        self.create_context()
        self.create_docx_images()
        self.split_items_to_chunks()
        self.render()

    def create_docx_instances(self):
        self.listing_docx = DocxTemplate(self.listing_template_path)
        self.gallery_docx = DocxTemplate(self.gallery_template_path)

    def create_context(self):
        self.context = { "items" : self.evidence_items,
                         "evidence": self.evidence,
                         "start_date": self.start_date,
                         "end_date": self.end_date,
                         "count": self.count,
                         "chunks": [],}

    def sort_items_by_date(self):
        if self.evidence_items and len(self.evidence_items) > 0:
            self.evidence_items
            items = sorted(
                    [item for item in self.evidence_items if (item["updated_at"] and isinstance(item["updated_at"], datetime))],
                    key=lambda item:(item["updated_at"]))

            undated_items = [item for item in self.evidence_items if not item["updated_at"]]
            items.extend(undated_items)
            self.evidence_items = items

    def create_docx_images(self):
        for item in self.context['items']:
            item["image"] = InlineImage(self.gallery_docx,
                                        image_descriptor=item["image_full_path"],
                                        height=Mm(40))

    def split_items_to_chunks(self):
        self.context["chunks"] = to_chunks(self.evidence_items, 3)

    def render(self):
        self.listing_docx.render(self.context)
        self.listing_docx.save(f"{EXPORT_PATH}/Evidence {self.evidence} list.docx")

        self.gallery_docx.render(self.context)
        self.gallery_docx.save(f"{EXPORT_PATH}/Evidence {self.evidence} images.docx")
