"""IDs de SAP GUI que usan las tareas.

Son IDs de ejemplo con la forma estándar de SAP GUI Scripting. Cada sistema puede
tener otras pantallas o subpantallas según su versión y configuración: grábalos en tu
sistema con el grabador de scripts de SAP GUI (Alt+F12 > Grabación y reproducción de
scripts) y reemplázalos aquí. Mantenerlos en un solo archivo hace que un cambio de
pantalla se corrija una sola vez.
"""

# Pantalla inicial de MM01 / MM02
MATERIAL = "wnd[0]/usr/ctxtRMMG1-MATNR"

# Ventana de selección de vistas (MM01 / MM02)
VISTAS_TABLA = "wnd[1]/usr/tblSAPLMGMMTC_VIEW"
VISTA_DATOS_BASE_1 = VISTAS_TABLA + "/txtMSICHTAUSW-DYTXT[0,0]"
VISTA_ALMACEN_1 = VISTAS_TABLA + "/txtMSICHTAUSW-DYTXT[0,13]"

# Ventana de niveles organizativos (MM01)
CENTRO = "wnd[1]/usr/ctxtRMMG1-WERKS"
ALMACEN = "wnd[1]/usr/ctxtRMMG1-LGORT"

# Vista Datos base 1 (MM02)
_DATOS_BASE = "wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004"
DESCRIPCION = _DATOS_BASE + "/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX"
STATUS_MATERIAL = _DATOS_BASE + "/subSUB2:SAPLMGD1:2001/ctxtMARA-MSTAE"
