from flask import Flask, request, redirect, render_template_string
from openpyxl import Workbook, load_workbook
from datetime import date
import os

# ============================================================
# CONFIGURACIÓN
# ============================================================

app = Flask(__name__)

ARCHIVO_EXCEL = "almacen_hospital.xlsx"


# ============================================================
# CREAR ARCHIVO EXCEL
# ============================================================

def crear_excel():

    if not os.path.exists(ARCHIVO_EXCEL):

        wb = Workbook()

        # HOJA PRODUCTOS
        ws_productos = wb.active
        ws_productos.title = "Productos"

        ws_productos.append([
            "ID",
            "Nombre",
            "Categoria",
            "Stock",
            "Stock Minimo"
        ])

        ws_productos.append([
            101,
            "Paracetamol 500mg",
            "Medicamentos",
            100,
            30
        ])

        # HOJA MOVIMIENTOS
        ws_movimientos = wb.create_sheet("Movimientos")

        ws_movimientos.append([
            "ID Movimiento",
            "ID Producto",
            "Tipo",
            "Cantidad",
            "Fecha",
            "Motivo"
        ])

        wb.save(ARCHIVO_EXCEL)


# ============================================================
# OBTENER PRODUCTOS
# ============================================================

def obtener_productos():

    crear_excel()

    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Productos"]

    productos = []

    for fila in range(2, ws.max_row + 1):

        if ws.cell(fila, 1).value is not None:

            productos.append({
                "id": ws.cell(fila, 1).value,
                "nombre": ws.cell(fila, 2).value,
                "categoria": ws.cell(fila, 3).value,
                "stock": ws.cell(fila, 4).value,
                "minimo": ws.cell(fila, 5).value
            })

    wb.close()

    return productos


# ============================================================
# GUARDAR PRODUCTO
# ============================================================

def guardar_producto(
    id_producto,
    nombre,
    categoria,
    stock,
    minimo
):

    crear_excel()

    wb = load_workbook(ARCHIVO_EXCEL)
    ws = wb["Productos"]

    # Verificar ID repetido
    for fila in range(2, ws.max_row + 1):

        if ws.cell(fila, 1).value == id_producto:

            wb.close()
            return False

    ws.append([
        id_producto,
        nombre,
        categoria,
        stock,
        minimo
    ])

    wb.save(ARCHIVO_EXCEL)
    wb.close()

    return True


# ============================================================
# REGISTRAR MOVIMIENTO
# ============================================================

def registrar_movimiento(
    id_producto,
    tipo,
    cantidad,
    motivo
):

    crear_excel()

    wb = load_workbook(ARCHIVO_EXCEL)

    ws_productos = wb["Productos"]
    ws_movimientos = wb["Movimientos"]

    fila_producto = None

    # Buscar producto
    for fila in range(2, ws_productos.max_row + 1):

        if ws_productos.cell(fila, 1).value == id_producto:

            fila_producto = fila
            break

    if fila_producto is None:

        wb.close()
        return False, "Producto no encontrado."

    stock_actual = ws_productos.cell(
        fila_producto,
        4
    ).value

    # ENTRADA
    if tipo == "Entrada":

        nuevo_stock = stock_actual + cantidad

    # SALIDA
    elif tipo == "Salida":

        if cantidad > stock_actual:

            wb.close()
            return False, "Stock insuficiente."

        nuevo_stock = stock_actual - cantidad

    else:

        wb.close()
        return False, "Tipo de movimiento inválido."

    # Actualizar stock
    ws_productos.cell(
        fila_producto,
        4
    ).value = nuevo_stock

    # Obtener último ID
    ultimo_id = 0

    for fila in range(2, ws_movimientos.max_row + 1):

        valor = ws_movimientos.cell(
            fila,
            1
        ).value

        if valor is not None:

            if int(valor) > ultimo_id:

                ultimo_id = int(valor)

    nuevo_id = ultimo_id + 1

    # Guardar movimiento
    ws_movimientos.append([
        nuevo_id,
        id_producto,
        tipo,
        cantidad,
        date.today(),
        motivo
    ])

    wb.save(ARCHIVO_EXCEL)
    wb.close()

    return True, "Movimiento registrado correctamente."


# ============================================================
# OBTENER MOVIMIENTOS
# ============================================================

def obtener_movimientos():

    crear_excel()

    wb = load_workbook(ARCHIVO_EXCEL)

    ws = wb["Movimientos"]

    movimientos = []

    for fila in range(2, ws.max_row + 1):

        if ws.cell(fila, 1).value is not None:

            movimientos.append({
                "id": ws.cell(fila, 1).value,
                "producto": ws.cell(fila, 2).value,
                "tipo": ws.cell(fila, 3).value,
                "cantidad": ws.cell(fila, 4).value,
                "fecha": ws.cell(fila, 5).value,
                "motivo": ws.cell(fila, 6).value
            })

    wb.close()

    movimientos.reverse()

    return movimientos


# ============================================================
# DISEÑO PRINCIPAL
# ============================================================

HTML = """

<!DOCTYPE html>

<html lang="es">

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width, initial-scale=1.0">

<title>Sistema de Gestión de Almacén</title>

<style>

/* ==========================================================
   CONFIGURACIÓN GENERAL
   ========================================================== */

* {
    box-sizing: border-box;
}

body {

    margin: 0;

    font-family:
    "Segoe UI",
    Arial,
    sans-serif;

    background: #f4f7fb;

    color: #1e293b;
}


/* ==========================================================
   ESTRUCTURA
   ========================================================== */

.layout {

    display: flex;

    min-height: 100vh;
}


/* ==========================================================
   MENÚ LATERAL
   ========================================================== */

.sidebar {

    width: 260px;

    background: #123b5d;

    color: white;

    position: fixed;

    left: 0;

    top: 0;

    bottom: 0;

    padding: 22px 15px;

    overflow-y: auto;

    z-index: 1000;
}


/* LOGO */

.logo {

    display: flex;

    align-items: center;

    gap: 12px;

    padding: 10px 8px 25px;

    border-bottom:
    1px solid
    rgba(255,255,255,0.15);

    margin-bottom: 20px;
}

.logo-icon {

    width: 45px;

    height: 45px;

    background: white;

    color: #123b5d;

    border-radius: 12px;

    display: flex;

    align-items: center;

    justify-content: center;

    font-size: 24px;
}

.logo-text strong {

    display: block;

    font-size: 15px;
}

.logo-text span {

    font-size: 11px;

    color: #cbd5e1;
}


/* TÍTULO DEL MENÚ */

.menu-title {

    font-size: 11px;

    text-transform: uppercase;

    color: #94a3b8;

    padding: 0 12px;

    margin: 20px 0 8px;
}


/* ENLACES */

.sidebar a {

    display: flex;

    align-items: center;

    gap: 12px;

    text-decoration: none;

    color: #dbeafe;

    padding: 13px 14px;

    margin: 5px 0;

    border-radius: 9px;

    transition: 0.2s;

    font-size: 14px;
}

.sidebar a:hover {

    background: rgba(255,255,255,0.12);

    color: white;

    transform: translateX(3px);
}

.sidebar a.active {

    background: white;

    color: #123b5d;

    font-weight: 600;
}

.menu-icon {

    width: 25px;

    text-align: center;

    font-size: 18px;
}


/* ==========================================================
   CONTENIDO
   ========================================================== */

.main {

    margin-left: 260px;

    width: calc(100% - 260px);

    min-height: 100vh;
}


/* ==========================================================
   BARRA SUPERIOR
   ========================================================== */

.topbar {

    height: 75px;

    background: white;

    border-bottom: 1px solid #e2e8f0;

    display: flex;

    align-items: center;

    justify-content: space-between;

    padding: 0 30px;

    position: sticky;

    top: 0;

    z-index: 900;
}

.topbar-title {

    font-size: 20px;

    font-weight: 700;

    color: #123b5d;
}

.topbar-subtitle {

    font-size: 12px;

    color: #64748b;

    margin-top: 3px;
}


/* USUARIO */

.user {

    display: flex;

    align-items: center;

    gap: 10px;
}

.user-avatar {

    width: 38px;

    height: 38px;

    border-radius: 50%;

    background: #e0f2fe;

    color: #0369a1;

    display: flex;

    align-items: center;

    justify-content: center;

    font-weight: bold;
}

.user-info strong {

    display: block;

    font-size: 13px;
}

.user-info span {

    font-size: 11px;

    color: #64748b;
}


/* ==========================================================
   CONTENEDOR
   ========================================================== */

.container {

    width: 94%;

    max-width: 1400px;

    margin: 0 auto;

    padding: 30px 0 40px;
}


/* ==========================================================
   TÍTULOS
   ========================================================== */

.page-header {

    margin-bottom: 25px;
}

.page-header h1 {

    margin: 0;

    color: #123b5d;

    font-size: 25px;
}

.page-header p {

    margin: 7px 0 0;

    color: #64748b;

    font-size: 14px;
}


/* ==========================================================
   TARJETAS
   ========================================================== */

.cards {

    display: grid;

    grid-template-columns:
    repeat(4, 1fr);

    gap: 18px;

    margin-bottom: 28px;
}

.card {

    background: white;

    border-radius: 14px;

    padding: 22px;

    border: 1px solid #e2e8f0;

    box-shadow:
    0 4px 15px
    rgba(15, 23, 42, 0.05);

    transition: 0.2s;
}

.card:hover {

    transform: translateY(-2px);

    box-shadow:
    0 8px 20px
    rgba(15, 23, 42, 0.08);
}

.card-top {

    display: flex;

    align-items: center;

    justify-content: space-between;
}

.card-icon {

    width: 45px;

    height: 45px;

    border-radius: 11px;

    display: flex;

    align-items: center;

    justify-content: center;

    font-size: 21px;

    background: #e0f2fe;
}

.card h3 {

    font-size: 13px;

    color: #64748b;

    margin: 18px 0 6px;
}

.numero {

    font-size: 30px;

    font-weight: 700;

    color: #123b5d;
}


/* ==========================================================
   SECCIONES
   ========================================================== */

.section {

    background: white;

    border: 1px solid #e2e8f0;

    border-radius: 14px;

    padding: 22px;

    margin-bottom: 25px;

    box-shadow:
    0 4px 15px
    rgba(15, 23, 42, 0.04);
}

.section-header {

    display: flex;

    justify-content: space-between;

    align-items: center;

    margin-bottom: 15px;
}

.section-header h2 {

    margin: 0;

    font-size: 18px;

    color: #123b5d;
}


/* ==========================================================
   BOTONES
   ========================================================== */

.btn {

    display: inline-block;

    text-decoration: none;

    border: none;

    border-radius: 8px;

    padding: 11px 17px;

    font-size: 13px;

    font-weight: 600;

    cursor: pointer;

    transition: 0.2s;
}

.btn-primary {

    background: #123b5d;

    color: white;
}

.btn-primary:hover {

    background: #0c2f4b;

    transform: translateY(-1px);
}

.btn-secondary {

    background: #e0f2fe;

    color: #075985;
}

.btn-secondary:hover {

    background: #bae6fd;
}


/* ==========================================================
   TABLAS
   ========================================================== */

.table-container {

    width: 100%;

    overflow-x: auto;
}

table {

    width: 100%;

    border-collapse: collapse;

    min-width: 650px;
}

th {

    background: #f1f5f9;

    color: #475569;

    padding: 14px;

    text-align: left;

    font-size: 12px;

    text-transform: uppercase;

    border-bottom: 1px solid #e2e8f0;
}

td {

    padding: 14px;

    border-bottom: 1px solid #eef2f7;

    font-size: 13px;
}

tbody tr:hover {

    background: #f8fafc;
}


/* ==========================================================
   ESTADOS
   ========================================================== */

.badge {

    display: inline-block;

    padding: 6px 10px;

    border-radius: 20px;

    font-size: 11px;

    font-weight: 600;
}

.badge-success {

    background: #dcfce7;

    color: #166534;
}

.badge-danger {

    background: #fee2e2;

    color: #991b1b;
}

.badge-entry {

    background: #dcfce7;

    color: #166534;
}

.badge-exit {

    background: #fee2e2;

    color: #991b1b;
}


/* ==========================================================
   FORMULARIOS
   ========================================================== */

.form-card {

    background: white;

    border: 1px solid #e2e8f0;

    border-radius: 14px;

    padding: 28px;

    max-width: 800px;

    box-shadow:
    0 4px 15px
    rgba(15, 23, 42, 0.04);
}

.form-grid {

    display: grid;

    grid-template-columns:
    repeat(2, 1fr);

    gap: 20px;
}

.form-group {

    margin-bottom: 3px;
}

.form-group.full {

    grid-column: 1 / -1;
}

label {

    display: block;

    margin-bottom: 7px;

    font-size: 13px;

    font-weight: 600;

    color: #334155;
}

input,
select {

    width: 100%;

    padding: 12px 13px;

    border: 1px solid #cbd5e1;

    border-radius: 8px;

    background: white;

    color: #1e293b;

    font-size: 13px;

    outline: none;

    transition: 0.2s;
}

input:focus,
select:focus {

    border-color: #0284c7;

    box-shadow:
    0 0 0 3px
    rgba(2,132,199,0.10);
}

.form-actions {

    margin-top: 22px;

    display: flex;

    gap: 10px;
}


/* ==========================================================
   ALERTAS
   ========================================================== */

.alert {

    padding: 17px;

    border-radius: 10px;

    margin-bottom: 20px;

    font-size: 13px;
}

.alert-success {

    background: #dcfce7;

    color: #166534;

    border: 1px solid #bbf7d0;
}

.alert-error {

    background: #fee2e2;

    color: #991b1b;

    border: 1px solid #fecaca;
}


/* ==========================================================
   ALERTA DE STOCK
   ========================================================== */

.stock-alert {

    border-left: 4px solid #dc2626;

    background: #fff7f7;
}


/* ==========================================================
   ESTADO VACÍO
   ========================================================== */

.empty {

    text-align: center;

    padding: 45px 20px;

    color: #64748b;
}

.empty-icon {

    font-size: 42px;

    margin-bottom: 10px;
}


/* ==========================================================
   PIE
   ========================================================== */

.footer {

    text-align: center;

    padding: 25px;

    color: #94a3b8;

    font-size: 12px;
}


/* ==========================================================
   BOTÓN MÓVIL
   ========================================================== */

.mobile-menu {

    display: none;

    border: none;

    background: transparent;

    font-size: 25px;

    cursor: pointer;

    color: #123b5d;
}


/* ==========================================================
   RESPONSIVE TABLET
   ========================================================== */

@media (max-width: 1100px) {

    .cards {

        grid-template-columns:
        repeat(2, 1fr);
    }

}


/* ==========================================================
   RESPONSIVE CELULAR
   ========================================================== */

@media (max-width: 768px) {

    .sidebar {

        transform: translateX(-100%);

        transition: 0.3s;

        width: 250px;
    }

    .sidebar.open {

        transform: translateX(0);
    }

    .main {

        margin-left: 0;

        width: 100%;
    }

    .topbar {

        padding: 0 18px;

        height: 65px;
    }

    .mobile-menu {

        display: block;
    }

    .topbar-title {

        font-size: 16px;
    }

    .topbar-subtitle {

        display: none;
    }

    .user-info {

        display: none;
    }

    .container {

        width: 92%;

        padding-top: 22px;
    }

    .cards {

        grid-template-columns: 1fr;

        gap: 12px;
    }

    .form-grid {

        grid-template-columns: 1fr;
    }

    .form-group.full {

        grid-column: auto;
    }

    .section {

        padding: 17px;
    }

    .form-card {

        padding: 20px;
    }

}


/* ==========================================================
   CELULAR PEQUEÑO
   ========================================================== */

@media (max-width: 450px) {

    .topbar {

        padding: 0 12px;
    }

    .page-header h1 {

        font-size: 21px;
    }

    .numero {

        font-size: 26px;
    }

}

</style>

</head>


<body>


<div class="layout">


<!-- ========================================================
     MENÚ LATERAL
     ======================================================== -->

<aside class="sidebar" id="sidebar">

    <div class="logo">

        <div class="logo-icon">
            🏥
        </div>

        <div class="logo-text">

            <strong>
                HOSPITAL
            </strong>

            <span>
                Simón Bolívar
            </span>

        </div>

    </div>


    <div class="menu-title">
        Principal
    </div>


    <a href="/" class="active">

        <span class="menu-icon">
            🏠
        </span>

        Inicio

    </a>


    <div class="menu-title">
        Inventario
    </div>


    <a href="/productos">

        <span class="menu-icon">
            📦
        </span>

        Productos

    </a>


    <a href="/nuevo-producto">

        <span class="menu-icon">
            ➕
        </span>

        Nuevo producto

    </a>


    <a href="/movimiento">

        <span class="menu-icon">
            🔄
        </span>

        Registrar movimiento

    </a>


    <div class="menu-title">
        Consultas
    </div>


    <a href="/movimientos">

        <span class="menu-icon">
            📋
        </span>

        Historial

    </a>


    <a href="/alertas">

        <span class="menu-icon">
            ⚠️
        </span>

        Alertas

    </a>


</aside>


<!-- ========================================================
     CONTENIDO PRINCIPAL
     ======================================================== -->

<main class="main">


    <!-- BARRA SUPERIOR -->

    <header class="topbar">

        <div style="display:flex;align-items:center;gap:15px;">

            <button
                class="mobile-menu"
                onclick="toggleMenu()"
            >
                ☰
            </button>

            <div>

                <div class="topbar-title">
                    Sistema de Gestión de Almacén
                </div>

                <div class="topbar-subtitle">
                    Control de inventario hospitalario
                </div>

            </div>

        </div>


        <div class="user">

            <div class="user-avatar">
                AL
            </div>

            <div class="user-info">

                <strong>
                    Almacén
                </strong>

                <span>
                    Encargado de almacén
                </span>

            </div>

        </div>

    </header>


    <!-- CONTENIDO -->

    <div class="container">

        {{ contenido | safe }}

    </div>


    <!-- PIE -->

    <div class="footer">

        Sistema de Gestión de Almacén
        ·
        Hospital Simón Bolívar

        <br>

        Gestión de inventario hospitalario

    </div>


</main>


</div>


<script>

function toggleMenu() {

    const sidebar =
        document.getElementById("sidebar");

    sidebar.classList.toggle("open");

}

</script>


</body>

</html>

"""


# ============================================================
# MOSTRAR PÁGINA
# ============================================================

def pagina(contenido):

    return render_template_string(
        HTML,
        contenido=contenido
    )


# ============================================================
# INICIO / DASHBOARD
# ============================================================

@app.route("/")
def inicio():

    productos = obtener_productos()

    total_productos = len(productos)

    stock_critico = 0

    total_stock = 0

    for producto in productos:

        total_stock += producto["stock"]

        if producto["stock"] <= producto["minimo"]:

            stock_critico += 1

    movimientos = obtener_movimientos()

    total_movimientos = len(movimientos)


    contenido = f"""

    <div class="page-header">

        <h1>
            Panel principal
        </h1>

        <p>
            Resumen general del inventario del hospital.
        </p>

    </div>


    <!-- TARJETAS -->

    <div class="cards">


        <div class="card">

            <div class="card-top">

                <div class="card-icon">
                    📦
                </div>

            </div>

            <h3>
                PRODUCTOS REGISTRADOS
            </h3>

            <div class="numero">
                {total_productos}
            </div>

        </div>


        <div class="card">

            <div class="card-top">

                <div class="card-icon">
                    📊
                </div>

            </div>

            <h3>
                STOCK TOTAL
            </h3>

            <div class="numero">
                {total_stock}
            </div>

        </div>


        <div class="card">

            <div class="card-top">

                <div class="card-icon">
                    🔄
                </div>

            </div>

            <h3>
                MOVIMIENTOS
            </h3>

            <div class="numero">
                {total_movimientos}
            </div>

        </div>


        <div class="card">

            <div class="card-top">

                <div class="card-icon">
                    ⚠️
                </div>

            </div>

            <h3>
                STOCK CRÍTICO
            </h3>

            <div class="numero">
                {stock_critico}
            </div>

        </div>


    </div>


    <!-- PRODUCTOS -->

    <div class="section">

        <div class="section-header">

            <h2>
                Productos registrados
            </h2>

            <a
                href="/nuevo-producto"
                class="btn btn-primary"
            >
                + Nuevo producto
            </a>

        </div>


        <div class="table-container">

        <table>

            <thead>

                <tr>

                    <th>ID</th>

                    <th>Producto</th>

                    <th>Categoría</th>

                    <th>Stock</th>

                    <th>Mínimo</th>

                    <th>Estado</th>

                </tr>

            </thead>

            <tbody>

    """


    for producto in productos:

        if producto["stock"] <= producto["minimo"]:

            estado = """

            <span class="badge badge-danger">
                ⚠️ Crítico
            </span>

            """

        else:

            estado = """

            <span class="badge badge-success">
                ✓ Estable
            </span>

            """


        contenido += f"""

                <tr>

                    <td>
                        <strong>
                            {producto["id"]}
                        </strong>
                    </td>

                    <td>
                        {producto["nombre"]}
                    </td>

                    <td>
                        {producto["categoria"]}
                    </td>

                    <td>
                        <strong>
                            {producto["stock"]}
                        </strong>
                    </td>

                    <td>
                        {producto["minimo"]}
                    </td>

                    <td>
                        {estado}
                    </td>

                </tr>

        """


    contenido += """

            </tbody>

        </table>

        </div>

    </div>

    """


    return pagina(contenido)


# ============================================================
# PRODUCTOS
# ============================================================

@app.route("/productos")
def productos():

    lista = obtener_productos()


    contenido = """

    <div class="page-header">

        <h1>
            📦 Productos
        </h1>

        <p>
            Consulta y control de productos almacenados.
        </p>

    </div>


    <div class="section">

        <div class="section-header">

            <h2>
                Inventario
            </h2>

            <a
                href="/nuevo-producto"
                class="btn btn-primary"
            >
                + Nuevo producto
            </a>

        </div>


        <div class="table-container">

        <table>

            <thead>

                <tr>

                    <th>ID</th>

                    <th>Producto</th>

                    <th>Categoría</th>

                    <th>Stock</th>

                    <th>Stock mínimo</th>

                    <th>Estado</th>

                </tr>

            </thead>

            <tbody>

    """


    if not lista:

        contenido += """

                <tr>

                    <td colspan="6">

                        <div class="empty">

                            <div class="empty-icon">
                                📦
                            </div>

                            No hay productos registrados.

                        </div>

                    </td>

                </tr>

        """


    for producto in lista:

        if producto["stock"] <= producto["minimo"]:

            estado = """

            <span class="badge badge-danger">
                ⚠️ Stock crítico
            </span>

            """

        else:

            estado = """

            <span class="badge badge-success">
                ✓ Stock estable
            </span>

            """


        contenido += f"""

                <tr>

                    <td>
                        <strong>
                            {producto["id"]}
                        </strong>
                    </td>

                    <td>
                        {producto["nombre"]}
                    </td>

                    <td>
                        {producto["categoria"]}
                    </td>

                    <td>
                        <strong>
                            {producto["stock"]}
                        </strong>
                    </td>

                    <td>
                        {producto["minimo"]}
                    </td>

                    <td>
                        {estado}
                    </td>

                </tr>

        """


    contenido += """

            </tbody>

        </table>

        </div>

    </div>

    """


    return pagina(contenido)


# ============================================================
# NUEVO PRODUCTO
# ============================================================

@app.route("/nuevo-producto")
def nuevo_producto():

    contenido = """

    <div class="page-header">

        <h1>
            ➕ Nuevo producto
        </h1>

        <p>
            Registra un nuevo producto en el inventario.
        </p>

    </div>


    <div class="form-card">

        <form
            action="/guardar-producto"
            method="POST"
        >

            <div class="form-grid">


                <div class="form-group">

                    <label>
                        ID del producto
                    </label>

                    <input
                        type="number"
                        name="id"
                        placeholder="Ejemplo: 102"
                        min="1"
                        required
                    >

                </div>


                <div class="form-group">

                    <label>
                        Nombre del producto
                    </label>

                    <input
                        type="text"
                        name="nombre"
                        placeholder="Ejemplo: Ibuprofeno 400mg"
                        required
                    >

                </div>


                <div class="form-group">

                    <label>
                        Categoría
                    </label>

                    <input
                        type="text"
                        name="categoria"
                        placeholder="Ejemplo: Medicamentos"
                        required
                    >

                </div>


                <div class="form-group">

                    <label>
                        Stock inicial
                    </label>

                    <input
                        type="number"
                        name="stock"
                        placeholder="Cantidad"
                        min="0"
                        required
                    >

                </div>


                <div class="form-group">

                    <label>
                        Stock mínimo
                    </label>

                    <input
                        type="number"
                        name="minimo"
                        placeholder="Cantidad mínima"
                        min="0"
                        required
                    >

                </div>


            </div>


            <div class="form-actions">

                <button
                    type="submit"
                    class="btn btn-primary"
                >
                    ✓ Guardar producto
                </button>

                <a
                    href="/productos"
                    class="btn btn-secondary"
                >
                    Cancelar
                </a>

            </div>


        </form>

    </div>

    """


    return pagina(contenido)


# ============================================================
# GUARDAR PRODUCTO
# ============================================================

@app.route(
    "/guardar-producto",
    methods=["POST"]
)
def guardar_nuevo_producto():

    try:

        id_producto = int(
            request.form["id"]
        )

        nombre = request.form["nombre"].strip()

        categoria = request.form["categoria"].strip()

        stock = int(
            request.form["stock"]
        )

        minimo = int(
            request.form["minimo"]
        )

    except (ValueError, KeyError):

        contenido = """

        <div class="alert alert-error">

            ❌ Los datos ingresados no son válidos.

            <br><br>

            <a
                href="/nuevo-producto"
                class="btn btn-secondary"
            >
                Volver
            </a>

        </div>

        """

        return pagina(contenido)


    if stock < 0 or minimo < 0:

        contenido = """

        <div class="alert alert-error">

            ❌ El stock no puede ser negativo.

            <br><br>

            <a
                href="/nuevo-producto"
                class="btn btn-secondary"
            >
                Volver
            </a>

        </div>

        """

        return pagina(contenido)


    resultado = guardar_producto(
        id_producto,
        nombre,
        categoria,
        stock,
        minimo
    )


    if resultado:

        return redirect("/productos")


    contenido = """

    <div class="alert alert-error">

        ❌ <strong>Producto duplicado</strong>

        <br><br>

        El ID del producto ya existe
        en el inventario.

        <br><br>

        <a
            href="/nuevo-producto"
            class="btn btn-secondary"
        >
            Volver
        </a>

    </div>

    """


    return pagina(contenido)


# ============================================================
# FORMULARIO MOVIMIENTO
# ============================================================

@app.route("/movimiento")
def movimiento():

    productos = obtener_productos()


    contenido = """

    <div class="page-header">

        <h1>
            🔄 Registrar movimiento
        </h1>

        <p>
            Registra entradas y salidas de productos.
        </p>

    </div>


    <div class="form-card">

        <form
            action="/guardar-movimiento"
            method="POST"
        >

            <div class="form-grid">


                <div class="form-group full">

                    <label>
                        Producto
                    </label>

                    <select
                        name="id_producto"
                        required
                    >

                        <option value="">
                            Seleccionar producto
                        </option>

    """


    if not productos:

        contenido += """

                        <option disabled>
                            No existen productos registrados
                        </option>

        """


    for producto in productos:

        contenido += f"""

                        <option
                            value="{producto["id"]}"
                        >

                            {producto["id"]}
                            -
                            {producto["nombre"]}
                            | Stock:
                            {producto["stock"]}

                        </option>

        """


    contenido += """

                    </select>

                </div>


                <div class="form-group">

                    <label>
                        Tipo de movimiento
                    </label>

                    <select
                        name="tipo"
                        required
                    >

                        <option value="Entrada">
                            📥 Entrada de stock
                        </option>

                        <option value="Salida">
                            📤 Salida de stock
                        </option>

                    </select>

                </div>


                <div class="form-group">

                    <label>
                        Cantidad
                    </label>

                    <input
                        type="number"
                        name="cantidad"
                        min="1"
                        placeholder="Cantidad"
                        required
                    >

                </div>


                <div class="form-group full">

                    <label>
                        Motivo
                    </label>

                    <input
                        type="text"
                        name="motivo"
                        placeholder="Ejemplo: Compra, distribución, reposición..."
                        required
                    >

                </div>


            </div>


            <div class="form-actions">

                <button
                    type="submit"
                    class="btn btn-primary"
                >
                    ✓ Registrar movimiento
                </button>

                <a
                    href="/"
                    class="btn btn-secondary"
                >
                    Cancelar
                </a>

            </div>


        </form>

    </div>

    """
    return pagina(contenido)

# ============================================================
# GUARDAR MOVIMIENTO
# ============================================================

@app.route(
    "/guardar-movimiento",
    methods=["POST"]
)
def guardar_nuevo_movimiento():

    try:

        id_producto = int(
            request.form["id_producto"]
        )

        tipo = request.form["tipo"]

        cantidad = int(
            request.form["cantidad"]
        )

        motivo = request.form["motivo"].strip()

    except (ValueError, KeyError):

        contenido = """

        <div class="alert alert-error">

            ❌ Los datos ingresados no son válidos.

            <br><br>

            <a
                href="/movimiento"
                class="btn btn-secondary"
            >
                Volver
            </a>

        </div>

        """

        return pagina(contenido)


    if cantidad <= 0:

        contenido = """

        <div class="alert alert-error">

            ❌ La cantidad debe ser mayor que cero.

            <br><br>

            <a
                href="/movimiento"
                class="btn btn-secondary"
            >
                Volver
            </a>

        </div>

        """

        return pagina(contenido)


    resultado, mensaje = registrar_movimiento(
        id_producto,
        tipo,
        cantidad,
        motivo
    )


    if resultado:

        return redirect("/productos")


    contenido = f"""

    <div class="alert alert-error">

        ❌ <strong>Error</strong>

        <br><br>

        {mensaje}

        <br><br>

        <a
            href="/movimiento"
            class="btn btn-secondary"
        >
            Volver
        </a>

    </div>

    """


    return pagina(contenido)


# ============================================================
# HISTORIAL
# ============================================================

@app.route("/movimientos")
def movimientos():

    lista = obtener_movimientos()


    contenido = """

    <div class="page-header">

        <h1>
            📋 Historial de movimientos
        </h1>

        <p>
            Registro de entradas y salidas del inventario.
        </p>

    </div>


    <div class="section">

        <div class="table-container">

        <table>

            <thead>

                <tr>

                    <th>ID</th>

                    <th>Producto</th>

                    <th>Tipo</th>

                    <th>Cantidad</th>

                    <th>Fecha</th>

                    <th>Motivo</th>

                </tr>

            </thead>

            <tbody>

    """


    if not lista:

        contenido += """

                <tr>

                    <td colspan="6">

                        <div class="empty">

                            <div class="empty-icon">
                                📋
                            </div>

                            Todavía no existen movimientos registrados.

                        </div>

                    </td>

                </tr>

        """


    for movimiento in lista:

        if movimiento["tipo"] == "Entrada":

            tipo = """

            <span class="badge badge-entry">
                📥 Entrada
            </span>

            """

        else:

            tipo = """

            <span class="badge badge-exit">
                📤 Salida
            </span>

            """


        contenido += f"""

                <tr>

                    <td>
                        <strong>
                            #{movimiento["id"]}
                        </strong>
                    </td>

                    <td>
                        ID {movimiento["producto"]}
                    </td>

                    <td>
                        {tipo}
                    </td>

                    <td>
                        <strong>
                            {movimiento["cantidad"]}
                        </strong>
                    </td>

                    <td>
                        {movimiento["fecha"]}
                    </td>

                    <td>
                        {movimiento["motivo"]}
                    </td>

                </tr>

        """


    contenido += """

            </tbody>

        </table>

        </div>

    </div>

    """


    return pagina(contenido)


# ============================================================
# ALERTAS
# ============================================================

@app.route("/alertas")
def alertas():

    productos = obtener_productos()

    productos_criticos = []


    for producto in productos:

        if producto["stock"] <= producto["minimo"]:

            productos_criticos.append(producto)


    contenido = """

    <div class="page-header">

        <h1>
            ⚠️ Alertas de inventario
        </h1>

        <p>
            Productos que necesitan revisión o reposición.
        </p>

    </div>

    """


    if not productos_criticos:

        contenido += """

        <div class="alert alert-success">

            ✓ <strong>Inventario estable</strong>

            <br><br>

            Ningún producto se encuentra
            por debajo del stock mínimo.

        </div>

        """


    else:

        contenido += f"""

        <div class="alert alert-error">

            ⚠️ Se encontraron
            <strong>
                {len(productos_criticos)}
            </strong>
            producto(s) con stock crítico.

        </div>


        <div class="section stock-alert">

            <div class="table-container">

            <table>

                <thead>

                    <tr>

                        <th>ID</th>

                        <th>Producto</th>

                        <th>Stock actual</th>

                        <th>Stock mínimo</th>

                        <th>Estado</th>

                    </tr>

                </thead>

                <tbody>

        """


        for producto in productos_criticos:

            contenido += f"""

                    <tr>

                        <td>
                            <strong>
                                {producto["id"]}
                            </strong>
                        </td>

                        <td>
                            {producto["nombre"]}
                        </td>

                        <td>
                            <strong>
                                {producto["stock"]}
                            </strong>
                        </td>

                        <td>
                            {producto["minimo"]}
                        </td>

                        <td>

                            <span class="badge badge-danger">
                                ⚠️ Requiere reposición
                            </span>

                        </td>

                    </tr>

            """


        contenido += """

                </tbody>

            </table>

            </div>

        </div>

        """


    return pagina(contenido)


# ============================================================
# EJECUTAR SERVIDOR
# ============================================================

if __name__ == "__main__":

    crear_excel()

    print("")
    print("==============================================")
    print(" HOSPITAL SIMÓN BOLÍVAR")
    print(" SISTEMA DE GESTIÓN DE ALMACÉN")
    print("==============================================")
    print("")
    print("Servidor iniciado.")
    print("")
    print("Abre tu navegador en:")
    print("http://127.0.0.1:5000")
    print("")
    print("El archivo Excel es:")
    print("almacen_hospital.xlsx")
    print("")
    print("Para detener el servidor:")
    print("Presiona CTRL + C")
    print("==============================================")
    print("")

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )