from pptx.shapes.graphfrm import GraphicFrame
from pptx.table import Table


class Style:
    def __init__(self, guid):
        self.guid = guid

    def apply_to(self, target):
        if isinstance(target, GraphicFrame) and hasattr(target, "table"):
            graphic_frame = target
        elif isinstance(target, Table):
            graphic_frame = target._graphic_frame
        else:
            raise TypeError("Expected a pptx Table, got %s instead." % target)
        graphic_frame._element.graphic.graphicData.tbl[0][-1].text = self.guid


class NoStyle:
    NoGrid = Style("{2D5ABB26-0587-4C30-8999-92F81FD0307C}")
    TableGrid = Style("{5940675A-B579-460E-94D1-54222C63F5DA}")


class ThemedStyle1:
    Accent1 = Style("{3C2FFA5D-87B4-456A-9821-1D50468CF0F}")
    Accent2 = Style("{284E427A-3D55-4303-BF80-6455036E1DE7}")
    Accent3 = Style("{69C7853C-536D-4A76-A0AE-DD22124D55A5}")
    Accent4 = Style("{775DCB02-9BB8-47FD-8907-85C794F793BA}")
    Accent5 = Style("{35758FB7-9AC5-4552-8A53-C91805E547FA}")
    Accent6 = Style("{08FB837D-C827-4EFA-A057-4D05807E0F7C}")


class ThemedStyle2:
    Accent1 = Style("{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}")
    Accent2 = Style("{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}")
    Accent3 = Style("{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}")
    Accent4 = Style("{E269D01E-BC32-4049-B463-5C60D7B0CCD2}")
    Accent5 = Style("{327F97BB-C833-4FB7-BDE5-3F7075034690}")
    Accent6 = Style("{638B1855-1B75-4FBE-930C-398BA8C253C6}")


class LightStyle1:
    Default = Style("{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}")
    Accent1 = Style("{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}")
    Accent2 = Style("{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}")
    Accent3 = Style("{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}")
    Accent4 = Style("{D27102A9-8310-4765-A935-A1911B00CA55}")
    Accent5 = Style("{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}")
    Accent6 = Style("{68D230F3-CF80-4859-8CE7-A43EE81993B5}")


class LightStyle2:
    Default = Style("{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}")
    Accent1 = Style("{69012ECD-51FC-41F1-AA8D-1B2483CD663E}")
    Accent2 = Style("{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}")
    Accent3 = Style("{F2DE63D5-997A-4646-A377-4702673A728D}")
    Accent4 = Style("{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}")
    Accent5 = Style("{5A111915-BE36-4E01-A7E5-04B1672EAD32}")
    Accent6 = Style("{912C8C85-51F0-491E-9774-3900AFEF0FD7}")


class LightStyle3:
    Default = Style("{616DA210-FB5B-4158-B5E0-FEB733F419BA}")
    Accent1 = Style("{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}")
    Accent2 = Style("{5DA37D80-6434-44D0-A028-1B22A696006F}")
    Accent3 = Style("{8799B23B-EC83-4686-B30A-512413B5E67A}")
    Accent4 = Style("{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}")
    Accent5 = Style("{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}")
    Accent6 = Style("{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}")


class MediumStyle1:
    Default = Style("{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}")
    Accent1 = Style("{B301B821-A1FF-4177-AEE7-76D212191A09}")
    Accent2 = Style("{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}")
    Accent3 = Style("{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}")
    Accent4 = Style("{1E171933-4619-4E11-9A3F-F7608DF75F80}")
    Accent5 = Style("{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}")
    Accent6 = Style("{10A1B5D5-9B99-4C35-A422-299274C87663}")


class MediumStyle2:
    Default = Style("{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}")
    Accent1 = Style("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
    Accent2 = Style("{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}")
    Accent3 = Style("{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}")
    Accent4 = Style("{00A15C55-8517-42AA-B614-E9B94910E393}")
    Accent5 = Style("{7DF18680-E054-41AD-8BC1-D1AEF772440D}")
    Accent6 = Style("{93296810-A885-4BE3-A3E7-6D5BEEA58F35}")


class MediumStyle3:
    Default = Style("{8EC20E35-A176-4012-BC5E-935CFFF8708E}")
    Accent1 = Style("{6E25E649-3F16-4E02-A733-19D2CDBF48F0}")
    Accent2 = Style("{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}")
    Accent3 = Style("{EB344D84-9AFB-497E-A393-DC336BA19D2E}")
    Accent4 = Style("{EB9631B5-78F2-41C9-869B-9F39066F8104}")
    Accent5 = Style("{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}")
    Accent6 = Style("{2A488322-F2BA-4B5B-9748-0D474271808F}")


class MediumStyle4:
    Default = Style("{D7AC3CCA-C797-4891-BE02-D94E43425B78}")
    Accent1 = Style("{69CF1AB2-1976-4502-BF36-3FF5EA218861}")
    Accent2 = Style("{8A107856-5554-42FB-B03E-39F5DBC370BA}")
    Accent3 = Style("{0505E3EF-67EA-436B-97B2-0124C06EBD24}")
    Accent4 = Style("{C4B1156A-380E-4F78-BDF5-A606A8083BF9}")
    Accent5 = Style("{22838BEF-8BB2-4498-84A7-C5851F593DF1}")
    Accent6 = Style("{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}")


class DarkStyle1:
    Default = Style("{E8034E78-7F5D-4C2E-B375-FC64B27BC917}")
    Accent1 = Style("{125E5076-3810-47DD-B79F-674D7AD40C01}")
    Accent2 = Style("{37CE84F3-28C3-443E-9E96-99CF82512B78}")
    Accent3 = Style("{D03447BB-5D67-496B-8E87-E561075AD55C}")
    Accent4 = Style("{E929F9F4-4A8F-4326-A1B4-22849713DDAB}")
    Accent5 = Style("{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}")
    Accent6 = Style("{AF606853-7671-496A-8E4F-DF71F8EC918B}")


class DarkStyle2:
    Default = Style("{5202B0CA-FC54-4496-8BCA-5EF66A818D29}")
    Accent1Accent2 = Style("{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}")
    Accent3Accent4 = Style("{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}")
    Accent5Accent6 = Style("{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}")
