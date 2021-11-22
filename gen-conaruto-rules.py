from odf import text, draw, teletype, table
from odf.opendocument import load
from unicodedata import normalize
from pathlib import Path
import json
doc = load("src/naruto-rules-template.odt")


def ssafe(s: any) -> str:
    """ Replace all specials chars in a string including accents

    :param any s: A string
    :return: The string with all specials char replaced
    """
    return normalize('NFD', str(s)).encode('ascii', 'ignore').decode("utf-8")

def values2str(values):
    v2s=[]
    for value in values:
        if value['name'] == "quantity":
            if'adjustement' in value and not value['adjustement'] == '=':
                v2s.append(f"{value['adjustement']}{value['count']}") 
            else:
                v2s.append(f"{value['count']}")
        elif value["name"] == "dice":
            if'adjustement' in value and not value['adjustement'] == '=':
                v2s.append(f"{value['count']}d{value['die']}")
            else:
                v2s.append(f"{value['count']}d{value['die']}")

    return ' '.join(v2s)
        

def property2str(obj, pname, ptype):
    if 'properties' in obj:
        oproperty = next(iter([p for p in obj['properties'] if p['name'] == pname and pname in p and p[pname] == ptype]), None)
        if oproperty is None:
            print(f"Object {obj['name']} has no property '{pname}'")
            return "-"
        else:
            if pname == "defense" and ptype == 'DEF':
                if 'values' in oproperty:
                    return ' '.join(values2str(oproperty['values'])) 
                else:
                    print(f"Object {obj['name']} has property '{pname}' with type '{ptype}' but no values")
                    return("-")
            elif pname == "fight" and ptype == "":
                if "damage" in oproperty and "values" in oproperty['damage']:
                    return ' '.join(values2str(oproperty['damage']['values'])) 
                else:
                    print(f"Object {obj['name']} has property '{pname}' with type '{ptype}' but no values")
                    return "-"
            elif pname == "measure" and ptype == "range":
                if "values" in oproperty:
                    return f"{' '.join(values2str(oproperty['values']))}{oproperty['unit']}"
                else:
                    print(f"Object {obj['name']} has property '{pname}' with type '{ptype}' but no values")
                    return "-"
            else:
                print(f"Object {obj['name']} has no property '{pname}' with type '{ptype}'")
                return "-"

    else:
        print(f"Object {obj['name']} has no properties")
        return "-"

def cost2str(cost):
    if cost is None:
        return "-"
    else:
        return f"{cost['value']} {cost['unit']}"
        
# for paragraph in doc.getElementsByType(text.P):
#     print(f" '{teletype.extractText(paragraph)}' : [{paragraph.getAttribute('stylename')}]")

start_tag = next(iter(p for p in doc.getElementsByType(text.P) if teletype.extractText(p) == "#start-tag#"), None)
end_tag = next(iter(p for p in doc.getElementsByType(text.P) if teletype.extractText(p) == "#end-tag#"), None)

pn = start_tag.parentNode

with open("src/naruto-ways.json") as nw:
    nws = json.load(nw)

    if start_tag is not None:
        profil = 0
        way_rank = 0
        print("We found start tag !")
        for item in nws:
            if item['otype'] == "capacity":
                #print(f"Add way rank '{item['name']}'...")
                way_rank += 1
                wr = text.P(stylename="WayRank")
                wr.addElement(text.Span(stylename="WayRankName", text=f"{way_rank}.\xa0{item['name']}"))
                if not item['voname'] == "Inconnu":
                    wr.addElement(text.Span(stylename="WayRankVoName", text=f" - {item['voname']} "))
                
                if item['limited']:
                    if not item['voname'] == "Inconnu":
                        wr.addElement(text.Span(stylename="WayRankVoName", text=f" "))
                    wr.addElement(text.Span(stylename="WayRankName", text="(L)\xa0:"))
                else:
                    wr.addElement(text.Span(stylename="WayRankName", text="\xa0:"))
                wr.addElement(text.Span(stylename="WayRankDescription", text=f" {item['full-description']}"))
                pn.insertBefore(wr,end_tag)
            elif item['otype'] == "way":
                print(f"Add way '{item['name']}'...")
                way_rank = 0
                wn = text.H(stylename="WayName", outlinelevel=3, text=f"{ssafe(item['name'])}")
                # wn = text.H(stylename="WayName", outlinelevel=3, text=f"{item['name']}")
                wd = text.P(stylename="WayDescription", text=f"{item['full-description']}")
                pn.insertBefore(wn,end_tag)
                pn.insertBefore(wd,end_tag)
            elif item['otype'] == "profil":
                profil += 1
                if profil > 5:
                    raise NotImplemented("You need to add more profil style ...")
                else:
                    print(f"Add profil '{item['name']}'...")
                    pds = text.P(stylename=f"ProfilDescriptionStart{profil}")
                    pd = text.P(stylename="ProfilDescription", text=f"{item['full-description']}")
                    pn.insertBefore(pds,end_tag)
                    pn.insertBefore(pd,end_tag)
                    for textbox in [h for h in doc.getElementsByType(text.H) if teletype.extractText(h).endswith(f"#profil-name-tag{profil}#")]:
                        print(f"Profiltag '{teletype.extractText(textbox)}' : [{textbox.getAttribute('stylename')}]")
                    profil_tag = next(iter(h for h in doc.getElementsByType(text.H) if teletype.extractText(h).endswith(f"#profil-name-tag{profil}#")), None)
                    if profil_tag is not None:
                        ptpn = profil_tag.parentNode
                        print("We found a profil name tag !")
                        pnh = text.H(stylename="ProfilName", outlinelevel=2, text=f"{item['name']}")           
                        ptpn.insertBefore(pnh,profil_tag)
                        ptpn.removeChild(profil_tag)
        
        pn.removeChild(start_tag)    
        if end_tag is not None:
            print("We found end tag !")   
            pn.removeChild(end_tag)

    with open("src/naruto-things.json") as no:
        nos = json.load(no)

        for ttype in ['Material', 'Armor', 'Weapon']:
            print(f"Object with no ttype {[o for o in nos if 'ttype' not in o]}")
            rtype = [ttype]
            if ttype == "Armor":
                rtype.append('Hat')
                
            # Find start and end tags
            otag_start = next(
                iter(
                    p for p in doc.getElementsByType(text.P) 
                    if teletype.extractText(p).startswith(f"#{ttype.lower()}-tag-start#")
                ), None
            )
            otag_end = next(
                iter(
                    p for p in doc.getElementsByType(text.P) 
                    if teletype.extractText(p).startswith(f"#{ttype.lower()}-tag-end#")
                ), None
            )

            if otag_start is None or otag_end is None:
                print(f"{ttype} tags not found !")
            else:
                print(f"{ttype} tags found ! ({otag_start} {otag_end})")
                opt = otag_start.parentNode
                for o in sorted([wo for wo in nos if wo['ttype'] in rtype], key = lambda i: i['name']):
                    on = text.P(stylename="WayRank")
                    on.addElement(text.Span(stylename="WayRankName", text=f"{o['name']}\xa0:"))
                    on.addElement(text.Span(stylename="WayRankDescription", text=f" {o['full-description']}"))
                    opt.insertBefore(on,otag_end)

                opt.removeChild(otag_start)    
                opt.removeChild(otag_end)

            otable_tag = next(
                iter(
                    p for p in doc.getElementsByType(text.P) 
                    if teletype.extractText(p).startswith(f"#{ttype.lower()}-name#")
                ), None
            )
            if otable_tag is None :
                print(f"TableCellTag not found !")
            else:
                print(f"TableCellTag '{teletype.extractText(otable_tag)}' : [{otable_tag.getAttribute('stylename')}]")
                otable_cell = otable_tag.parentNode
                otable_row = otable_cell.parentNode
                otable = otable_row.parentNode
                cell_style = otable_cell.getAttribute('stylename')
                row_style = otable_row.getAttribute('stylename')
                for o in sorted([wo for wo in nos if wo['ttype'] == ttype], key = lambda i: i['name']):
                    print(f"Add row : {o['name']}, {cost2str(o['cost'])}")
                    tr = table.TableRow(stylename=row_style)
                    otable.addElement(tr)
                    fc = table.TableCell(stylename=cell_style)
                    tr.addElement(fc)
                    fc.addElement(text.P(text=f"{o['name']}", stylename="RowCellFirst"))
                    if ttype == 'Weapon':
                        if 'range' not in o:
                            o.update({'range': None})
                        print(f"Add weapon cols : {property2str(o, 'fight', '')}, {property2str(o, 'measure', 'range')}")
                        c1 = table.TableCell(stylename=cell_style)
                        tr.addElement(c1)
                        c1.addElement(text.P(text=f"{property2str(o, 'fight', '')}", stylename="RowCell"))
                        c2 = table.TableCell(stylename=cell_style)
                        tr.addElement(c2)
                        c2.addElement(text.P(text=f"{property2str(o, 'measure', 'range')}", stylename="RowCell"))
                    if ttype == 'Armor' or ttype == 'Hat':
                        c1 = table.TableCell(stylename=cell_style)
                        tr.addElement(c1)
                        c1.addElement(text.P(text=f"{property2str(o,'defense','DEF')}", stylename="RowCell"))
                    c3 = table.TableCell(stylename=cell_style)
                    tr.addElement(c3)
                    c3.addElement(text.P(text=f"{cost2str(o['cost'])}", stylename="RowCell"))
                
                otable.removeChild(otable_row)

    gen_path = Path("generated")
    if not gen_path.exists() :
        gen_path.mkdir()
            
    doc.save("generated/naruto-rules-generated.odt")