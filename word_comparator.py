import io
import difflib
from docx import Document


def _extract_paragraphs(doc_bytes):
    doc = Document(io.BytesIO(doc_bytes))
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]


def _stats(paragraphs):
    texte = " ".join(paragraphs)
    return {
        "nb_paragraphes": len(paragraphs),
        "nb_mots": len(texte.split()),
        "nb_caracteres": len(texte),
    }


def compare_documents(doc1_bytes, doc2_bytes):
    """
    Compare deux documents Word (bytes) paragraphe par paragraphe.

    Retourne un dict avec les stats, les différences et un résumé textuel.
    """
    paras1 = _extract_paragraphs(doc1_bytes)
    paras2 = _extract_paragraphs(doc2_bytes)

    set1 = set(paras1)
    set2 = set(paras2)

    paragraphes_supprimes = [p for p in paras1 if p not in set2]
    paragraphes_ajoutes = [p for p in paras2 if p not in set1]

    # Paragraphes modifiés : paires proches détectées par difflib
    paragraphes_modifies = []
    matcher = difflib.SequenceMatcher(None, paras1, paras2)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "replace":
            for old, new in zip(paras1[i1:i2], paras2[j1:j2]):
                ratio = difflib.SequenceMatcher(None, old, new).ratio()
                if ratio < 1.0:
                    paragraphes_modifies.append({
                        "avant": old,
                        "apres": new,
                        "similarite": round(ratio, 2),
                    })

    # Taux de changement global
    total = max(len(paras1), len(paras2), 1)
    nb_changes = len(paragraphes_ajoutes) + len(paragraphes_supprimes) + len(paragraphes_modifies)
    taux_changement = round(min(nb_changes / total * 100, 100), 1)

    # Résumé textuel
    lignes_resume = []
    if not paragraphes_ajoutes and not paragraphes_supprimes and not paragraphes_modifies:
        lignes_resume.append("Les deux documents sont identiques.")
    else:
        if paragraphes_ajoutes:
            lignes_resume.append(f"{len(paragraphes_ajoutes)} paragraphe(s) ajouté(s) dans le document 2.")
        if paragraphes_supprimes:
            lignes_resume.append(f"{len(paragraphes_supprimes)} paragraphe(s) supprimé(s) par rapport au document 1.")
        if paragraphes_modifies:
            lignes_resume.append(f"{len(paragraphes_modifies)} paragraphe(s) modifié(s).")
        lignes_resume.append(f"Taux de changement global : {taux_changement}%.")

    return {
        "doc1_stats": _stats(paras1),
        "doc2_stats": _stats(paras2),
        "paragraphes_ajoutes": paragraphes_ajoutes,
        "paragraphes_supprimes": paragraphes_supprimes,
        "paragraphes_modifies": paragraphes_modifies,
        "resume_changements": " ".join(lignes_resume),
        "taux_changement": taux_changement,
    }
