"""
xml_parser.py
Lee el archivo stylesWordToDita.xml de OxygenAuthor y devuelve
la lista de estilos conocidos como una lista de diccionarios.
"""

import xml.etree.ElementTree as ET
import os


def parse_oxygen_styles(xml_path="data/stylesWordToDita.xml"):
    """
    Lee el XML de Oxygen y devuelve todos los estilos con nombre definido.
    Cada estilo tiene:
      - element_type : 'p' (párrafo) o 'r' (carácter/run)
      - style_name   : nombre del estilo en Word, ej. "Heading 1"
      - target_html  : elemento HTML al que se convierte, ej. "h1"
      - creates_new_block : True si genera un bloque nuevo (fresh="true")
    """
    if not os.path.exists(xml_path):
        print(f"⚠️  No se encontró el archivo XML en: {xml_path}")
        return []

    tree = ET.parse(xml_path)
    root = tree.getroot()

    styles = []
    to_html = root.find("toHTML")
    if to_html is None:
        return styles

    for relation in to_html.findall("relation"):
        element_raw  = (relation.findtext("element") or "").strip()
        style_name   = (relation.findtext("styleName") or "").strip()
        target_html  = (relation.findtext("resultedHTML") or "").strip()
        fresh_attr   = (relation.find("resultedHTML").get("fresh", "false")
                        if relation.find("resultedHTML") is not None else "false")

        # Solo nos interesan las relaciones que tienen nombre de estilo
        if not style_name:
            continue

        # El tipo de elemento es la parte antes del punto o los dos puntos
        # Ejemplos: "p", "r", "b", "i", "p.Body", "p:unordered-list(1)"
        element_type = element_raw.split(".")[0].split(":")[0]

        styles.append({
            "element_type":      element_type,
            "style_name":        style_name,
            "target_html":       target_html,
            "creates_new_block": fresh_attr.lower() == "true",
        })

    return styles


if __name__ == "__main__":
    # Prueba rápida: ejecuta `python xml_parser.py` para ver los estilos
    result = parse_oxygen_styles()
    print(f"Estilos encontrados: {len(result)}")
    for s in result[:10]:
        print(s)
