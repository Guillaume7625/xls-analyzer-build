# ====================================================================================
# PROVENANCE : DÉCOMPILATION PROPRE (pycdc) — constantes identiques au binaire,
# vérifiées contre SURFACE_SECURITE.md. Seules les bannières d'outillage ont été ôtées.
# ====================================================================================
'''Configuration canonique xls-analyzer : seuils chiffres + codes de rejet.
Source unique de verite. Tunables, mais valeurs V1 de production.'''
from __future__ import annotations
MAX_UPLOAD_BYTES = 26214400
MAX_TOTAL_UNCOMPRESSED = 209715200
MAX_ENTRY_UNCOMPRESSED = 104857600
MAX_ZIP_ENTRIES = 5000
MAX_COMPRESSION_RATIO = 100
MIN_ENTRY_BYTES_FOR_RATIO = 4096
MAX_SHAREDSTRINGS_BYTES = 52428800
MAX_XML_DEPTH = 100
RECALC_TIMEOUT_SECONDS = 120
RECALC_MEM_LIMIT = '512m'
RECALC_PIDS_LIMIT = 128
RECALC_CPUS = '1.0'
LOCK_STALE_SECONDS = 900
BIND_HOST = '127.0.0.1'
ENABLE_PUBLIC_DOCS = False
EVIDENCE_MAX_CHARS = 64
REQUIRED_OOXML_PARTS = ('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml')
ALLOWED_EXTENSIONS = ('.xlsx',)
