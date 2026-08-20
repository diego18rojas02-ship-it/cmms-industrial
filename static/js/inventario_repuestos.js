"use strict";

const modalFormulario = document.getElementById("modalFormulario");
const fondoFormulario = document.getElementById("fondoFormulario");
const cerrarFormulario = document.getElementById("cerrarFormulario");
const cancelarFormulario = document.getElementById("cancelarFormulario");
const guardarFormulario = document.getElementById("guardarFormulario");
const contenidoFormulario = document.getElementById("contenidoFormulario");
const modalTitulo = document.getElementById("modalTitulo");
const modalEtiqueta = document.getElementById("modalEtiqueta");
const errorFormulario = document.getElementById("errorFormulario");
const nuevaSeccion = document.getElementById("nuevaSeccion");
const agregarMaterial = document.getElementById("agregarMaterial");
const buscarInventario = document.getElementById("buscarInventario");
const filtrarEstadoStock = document.getElementById("filtrarEstadoStock");

let modoFormulario = "";
let idFormulario = null;
let seccionPadreFormulario = null;
let materialSeleccionado = null;
let temporizadorBusqueda = null;

window.INVENTARIO_SECCIONES =
    window.INVENTARIO_SECCIONES || [];

window.INVENTARIO_SUBSECCIONES =
    window.INVENTARIO_SUBSECCIONES || [];


function escaparHtml(texto) {
    return String(texto ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


function mostrarError(mensaje) {
    if (!errorFormulario) {
        alert(mensaje);
        return;
    }

    errorFormulario.textContent = mensaje;
    errorFormulario.classList.remove("oculto");
}


function limpiarError() {
    if (!errorFormulario) {
        return;
    }

    errorFormulario.textContent = "";
    errorFormulario.classList.add("oculto");
}


function abrirModal(etiqueta, titulo) {
    if (!modalFormulario) {
        return;
    }

    if (modalEtiqueta) {
        modalEtiqueta.textContent = etiqueta;
    }

    if (modalTitulo) {
        modalTitulo.textContent = titulo;
    }

    limpiarError();

    modalFormulario.classList.remove("oculto");
}


function cerrarModal() {
    if (!modalFormulario) {
        return;
    }

    modalFormulario.classList.add("oculto");

    if (contenidoFormulario) {
        contenidoFormulario.innerHTML = "";
    }

    limpiarError();

    modoFormulario = "";
    idFormulario = null;
    seccionPadreFormulario = null;
    materialSeleccionado = null;
}


function opcionesSecciones(seleccionada = "") {
    return window.INVENTARIO_SECCIONES
        .filter(
            item => Number(item.activo) === 1
        )
        .map(
            item => `
                <option
                    value="${item.id}"
                    ${
                        String(item.id) === String(seleccionada)
                            ? "selected"
                            : ""
                    }
                >
                    ${escaparHtml(item.nombre)}
                </option>
            `
        )
        .join("");
}


function opcionesSubsecciones(
    seccionId,
    seleccionada = ""
) {
    return window.INVENTARIO_SUBSECCIONES
        .filter(
            item =>
                Number(item.activo) === 1 &&
                Number(item.seccion_id) === Number(seccionId)
        )
        .map(
            item => `
                <option
                    value="${item.id}"
                    ${
                        String(item.id) === String(seleccionada)
                            ? "selected"
                            : ""
                    }
                >
                    ${escaparHtml(item.nombre)}
                </option>
            `
        )
        .join("");
}


function formularioSeccion(
    nombre = "",
    descripcion = ""
) {
    contenidoFormulario.innerHTML = `
        <div class="rejilla-formulario">

            <div class="campo campo-ancho">
                <label>Nombre</label>

                <input
                    type="text"
                    id="nombreEntidad"
                    value="${escaparHtml(nombre)}"
                    autocomplete="off"
                >
            </div>

            <div class="campo campo-ancho">
                <label>Descripción</label>

                <textarea
                    id="descripcionEntidad"
                    rows="3"
                >${escaparHtml(descripcion)}</textarea>
            </div>

        </div>
    `;

    setTimeout(() => {
        const campo = document.getElementById("nombreEntidad");

        if (campo) {
            campo.focus();
        }
    }, 50);
}


function formularioSubseccion(
    nombre = "",
    descripcion = ""
) {
    formularioSeccion(
        nombre,
        descripcion
    );
}


function formularioMaterial(datos = {}) {
    const seccionActual =
        datos.seccion_id || "";

    const subseccionActual =
        datos.subseccion_id || "";

    contenidoFormulario.innerHTML = `
        <div class="rejilla-formulario">

            ${
                modoFormulario === "crear_material"
                    ? `
                        <div class="campo campo-ancho">

                            <label>
                                Buscar material en SAP
                            </label>

                            <input
                                type="text"
                                id="buscarMaterialSap"
                                placeholder="Código SAP o descripción..."
                                autocomplete="off"
                            >

                            <div
                                id="resultadosMaterialSap"
                                class="resultados-busqueda"
                            ></div>

                        </div>

                        <div
                            id="materialSeleccionadoInfo"
                            class="material-inventario-seleccionado campo-ancho oculto"
                        ></div>
                    `
                    : ""
            }

            <div class="campo">

                <label>
                    Sección
                </label>

                <select id="seccionMaterial">

                    <option value="">
                        Seleccione
                    </option>

                    ${opcionesSecciones(seccionActual)}

                </select>

            </div>

            <div class="campo">

                <label>
                    Subsección
                </label>

                <select id="subseccionMaterial">

                    <option value="">
                        Sin subsección
                    </option>

                    ${opcionesSubsecciones(
                        seccionActual,
                        subseccionActual
                    )}

                </select>

            </div>

            <div class="campo">

                <label>
                    Stock mínimo
                </label>

                <input
                    type="number"
                    id="stockMinimoMaterial"
                    min="0"
                    step="any"
                    value="${datos.stock_minimo ?? 0}"
                >

            </div>

            <div class="campo">

                <label>
                    Stock objetivo
                </label>

                <input
                    type="number"
                    id="stockObjetivoMaterial"
                    min="0"
                    step="any"
                    value="${datos.stock_objetivo ?? 0}"
                >

            </div>

            <div class="campo campo-ancho">

                <label>
                    Observaciones
                </label>

                <textarea
                    id="observacionesMaterial"
                    rows="3"
                >${escaparHtml(
                    datos.observaciones || ""
                )}</textarea>

            </div>

        </div>
    `;

    const seccionMaterial =
        document.getElementById("seccionMaterial");

    const subseccionMaterial =
        document.getElementById("subseccionMaterial");

    if (
        seccionMaterial &&
        subseccionMaterial
    ) {
        seccionMaterial.addEventListener(
            "change",
            function () {
                subseccionMaterial.innerHTML = `
                    <option value="">
                        Sin subsección
                    </option>

                    ${opcionesSubsecciones(this.value)}
                `;
            }
        );
    }

    if (modoFormulario === "crear_material") {
        prepararBuscadorSap();
    }
}


function prepararBuscadorSap() {
    const campo =
        document.getElementById("buscarMaterialSap");

    const resultados =
        document.getElementById("resultadosMaterialSap");

    const info =
        document.getElementById("materialSeleccionadoInfo");

    if (
        !campo ||
        !resultados ||
        !info
    ) {
        return;
    }

    campo.addEventListener(
        "input",
        function () {
            clearTimeout(
                temporizadorBusqueda
            );

            const texto =
                this.value.trim();

            materialSeleccionado = null;

            info.innerHTML = "";
            info.classList.add("oculto");

            if (texto.length < 2) {
                resultados.innerHTML = "";
                return;
            }

            resultados.innerHTML = `
                <div class="resultado-cargando">
                    Buscando...
                </div>
            `;

            temporizadorBusqueda =
                setTimeout(
                    async () => {
                        try {
                            const respuesta =
                                await fetch(
                                    `/repuestos/api/materiales?q=${encodeURIComponent(
                                        texto
                                    )}`
                                );

                            if (!respuesta.ok) {
                                throw new Error(
                                    "No fue posible consultar la base SAP."
                                );
                            }

                            const materiales =
                                await respuesta.json();

                            if (!materiales.length) {
                                resultados.innerHTML = `
                                    <div class="sin-resultados">

                                        <strong>
                                            No se encontró en SAP
                                        </strong>

                                        <span>
                                            Puede crear el material manualmente.
                                        </span>

                                        <button
                                            type="button"
                                            class="boton-secundario"
                                            id="crearManualInventario"
                                        >
                                            Crear manual
                                        </button>

                                    </div>
                                `;

                                const botonManual =
                                    document.getElementById(
                                        "crearManualInventario"
                                    );

                                if (botonManual) {
                                    botonManual.addEventListener(
                                        "click",
                                        function () {
                                            crearMaterialManual(
                                                texto,
                                                info,
                                                resultados
                                            );
                                        }
                                    );
                                }

                                return;
                            }

                            resultados.innerHTML =
                                materiales
                                    .map(
                                        material => `
                                            <button
                                                type="button"
                                                class="resultado-material material-sap-inventario"
                                                data-id="${material.id}"
                                            >

                                                <div>

                                                    <strong>
                                                        ${escaparHtml(
                                                            material.codigo_sap
                                                            || "SIN CÓDIGO"
                                                        )}
                                                    </strong>

                                                    <span>
                                                        ${escaparHtml(
                                                            material.descripcion
                                                        )}
                                                    </span>

                                                </div>

                                                <div class="resultado-datos">

                                                    <span>
                                                        ${escaparHtml(
                                                            material.unidad_medida
                                                            || "-"
                                                        )}
                                                    </span>

                                                    <strong>
                                                        ${Number(
                                                            material.existencia
                                                            || 0
                                                        ).toLocaleString(
                                                            "es-CO"
                                                        )}
                                                    </strong>

                                                </div>

                                            </button>
                                        `
                                    )
                                    .join("");

                            document
                                .querySelectorAll(
                                    ".material-sap-inventario"
                                )
                                .forEach(
                                    boton => {
                                        boton.addEventListener(
                                            "click",
                                            function () {
                                                const id =
                                                    Number(
                                                        this.dataset.id
                                                    );

                                                materialSeleccionado =
                                                    materiales.find(
                                                        item =>
                                                            Number(item.id) === id
                                                    );

                                                if (!materialSeleccionado) {
                                                    return;
                                                }

                                                mostrarMaterialSeleccionado(
                                                    materialSeleccionado,
                                                    info
                                                );

                                                resultados.innerHTML = "";
                                            }
                                        );
                                    }
                                );

                        } catch (error) {
                            resultados.innerHTML = `
                                <div class="resultado-error">
                                    ${escaparHtml(error.message)}
                                </div>
                            `;
                        }
                    },
                    300
                );
        }
    );
}


function crearMaterialManual(
    textoBusqueda,
    info,
    resultados
) {
    const descripcion =
        prompt(
            "Descripción del material:",
            textoBusqueda
        );

    if (
        descripcion === null ||
        !descripcion.trim()
    ) {
        return;
    }

    const codigo =
        prompt(
            "Código SAP. Si no existe, escriba CREAR:",
            "CREAR"
        );

    if (codigo === null) {
        return;
    }

    const unidad =
        prompt(
            "Unidad de medida:",
            "UND"
        );

    if (
        unidad === null ||
        !unidad.trim()
    ) {
        return;
    }

    materialSeleccionado = {
        codigo_sap:
            codigo.trim() || "CREAR",

        descripcion:
            descripcion.trim(),

        unidad_medida:
            unidad.trim(),

        existencia:
            0,

        valor_unitario:
            0,

        origen_dato:
            "MANUAL"
    };

    mostrarMaterialSeleccionado(
        materialSeleccionado,
        info
    );

    resultados.innerHTML = "";
}


function mostrarMaterialSeleccionado(
    material,
    info
) {
    info.innerHTML = `
        <strong>
            ${escaparHtml(
                material.descripcion
            )}
        </strong>

        <span>
            ${escaparHtml(
                material.codigo_sap || "CREAR"
            )}
            ·
            ${escaparHtml(
                material.unidad_medida || "-"
            )}
            · Stock SAP:
            ${Number(
                material.existencia || 0
            ).toLocaleString("es-CO")}
        </span>
    `;

    info.classList.remove("oculto");
}


async function enviarJson(
    url,
    metodo,
    datos = null
) {
    const opciones = {
        method: metodo,
        headers: {
            "Content-Type":
                "application/json"
        }
    };

    if (datos !== null) {
        opciones.body =
            JSON.stringify(datos);
    }

    const respuesta =
        await fetch(
            url,
            opciones
        );

    let resultado;

    try {
        resultado =
            await respuesta.json();
    } catch (error) {
        throw new Error(
            "El servidor no devolvió una respuesta válida."
        );
    }

    if (
        !respuesta.ok ||
        !resultado.ok
    ) {
        throw new Error(
            resultado.error ||
            "La operación no pudo completarse."
        );
    }

    return resultado;
}


function valorCampo(id) {
    const elemento =
        document.getElementById(id);

    if (!elemento) {
        return "";
    }

    return elemento.value;
}


async function guardar() {
    limpiarError();

    guardarFormulario.disabled = true;
    guardarFormulario.textContent =
        "Guardando...";

    try {
        if (modoFormulario === "crear_seccion") {
            await enviarJson(
                "/repuestos/api/inventario/secciones",
                "POST",
                {
                    nombre:
                        valorCampo("nombreEntidad"),

                    descripcion:
                        valorCampo("descripcionEntidad")
                }
            );
        }

        else if (
            modoFormulario === "editar_seccion"
        ) {
            await enviarJson(
                `/repuestos/api/inventario/secciones/${idFormulario}`,
                "PUT",
                {
                    accion: "editar",

                    nombre:
                        valorCampo("nombreEntidad"),

                    descripcion:
                        valorCampo("descripcionEntidad")
                }
            );
        }

        else if (
            modoFormulario === "crear_subseccion"
        ) {
            await enviarJson(
                "/repuestos/api/inventario/subsecciones",
                "POST",
                {
                    seccion_id:
                        seccionPadreFormulario,

                    nombre:
                        valorCampo("nombreEntidad"),

                    descripcion:
                        valorCampo("descripcionEntidad")
                }
            );
        }

        else if (
            modoFormulario === "editar_subseccion"
        ) {
            await enviarJson(
                `/repuestos/api/inventario/subsecciones/${idFormulario}`,
                "PUT",
                {
                    accion: "editar",

                    nombre:
                        valorCampo("nombreEntidad"),

                    descripcion:
                        valorCampo("descripcionEntidad")
                }
            );
        }

        else if (
            modoFormulario === "crear_material"
        ) {
            if (!materialSeleccionado) {
                throw new Error(
                    "Seleccione un material SAP o créelo manualmente."
                );
            }

            const seccionId =
                valorCampo("seccionMaterial");

            if (!seccionId) {
                throw new Error(
                    "Seleccione una sección."
                );
            }

            const stockMinimo =
                Number(
                    valorCampo(
                        "stockMinimoMaterial"
                    )
                    || 0
                );

            const stockObjetivo =
                Number(
                    valorCampo(
                        "stockObjetivoMaterial"
                    )
                    || 0
                );

            if (
                !Number.isFinite(stockMinimo) ||
                stockMinimo < 0
            ) {
                throw new Error(
                    "El stock mínimo no es válido."
                );
            }

            if (
                !Number.isFinite(stockObjetivo) ||
                stockObjetivo < stockMinimo
            ) {
                throw new Error(
                    "El stock objetivo debe ser igual o mayor al stock mínimo."
                );
            }

            await enviarJson(
                "/repuestos/api/inventario/materiales",
                "POST",
                {
                    codigo_sap:
                        materialSeleccionado.codigo_sap,

                    descripcion:
                        materialSeleccionado.descripcion,

                    unidad_medida:
                        materialSeleccionado.unidad_medida,

                    seccion_id:
                        seccionId,

                    subseccion_id:
                        valorCampo(
                            "subseccionMaterial"
                        ),

                    stock_minimo:
                        stockMinimo,

                    stock_objetivo:
                        stockObjetivo,

                    observaciones:
                        valorCampo(
                            "observacionesMaterial"
                        ),

                    origen_dato:
                        materialSeleccionado.origen_dato
                        || "SAP"
                }
            );
        }

        else if (
            modoFormulario === "editar_material"
        ) {
            const seccionId =
                valorCampo("seccionMaterial");

            if (!seccionId) {
                throw new Error(
                    "Seleccione una sección."
                );
            }

            const stockMinimo =
                Number(
                    valorCampo(
                        "stockMinimoMaterial"
                    )
                    || 0
                );

            const stockObjetivo =
                Number(
                    valorCampo(
                        "stockObjetivoMaterial"
                    )
                    || 0
                );

            if (
                !Number.isFinite(stockMinimo) ||
                stockMinimo < 0
            ) {
                throw new Error(
                    "El stock mínimo no es válido."
                );
            }

            if (
                !Number.isFinite(stockObjetivo) ||
                stockObjetivo < stockMinimo
            ) {
                throw new Error(
                    "El stock objetivo debe ser igual o mayor al stock mínimo."
                );
            }

            await enviarJson(
                `/repuestos/api/inventario/materiales/${idFormulario}`,
                "PUT",
                {
                    accion:
                        "editar",

                    seccion_id:
                        seccionId,

                    subseccion_id:
                        valorCampo(
                            "subseccionMaterial"
                        ),

                    stock_minimo:
                        stockMinimo,

                    stock_objetivo:
                        stockObjetivo,

                    observaciones:
                        valorCampo(
                            "observacionesMaterial"
                        )
                }
            );
        }

        else {
            throw new Error(
                "No se identificó la operación que desea realizar."
            );
        }

        window.location.reload();

    } catch (error) {
        mostrarError(
            error.message
        );

    } finally {
        guardarFormulario.disabled =
            false;

        guardarFormulario.textContent =
            "Guardar";
    }
}


async function cambiarEstado(
    tipo,
    id,
    accion
) {
    try {
        await enviarJson(
            `/repuestos/api/inventario/${tipo}/${id}`,
            "PUT",
            {
                accion: accion
            }
        );

        window.location.reload();

    } catch (error) {
        alert(
            error.message
        );
    }
}


async function eliminarEntidad(
    tipo,
    id,
    mensaje
) {
    const confirmar =
        confirm(mensaje);

    if (!confirmar) {
        return;
    }

    try {
        await enviarJson(
            `/repuestos/api/inventario/${tipo}/${id}`,
            "DELETE"
        );

        window.location.reload();

    } catch (error) {
        alert(
            error.message
        );
    }
}


if (nuevaSeccion) {
    nuevaSeccion.addEventListener(
        "click",
        function () {
            modoFormulario =
                "crear_seccion";

            idFormulario =
                null;

            abrirModal(
                "SECCIÓN",
                "Nueva sección"
            );

            formularioSeccion();
        }
    );
}


document
    .querySelectorAll(".editar-seccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    modoFormulario =
                        "editar_seccion";

                    idFormulario =
                        Number(
                            this.dataset.id
                        );

                    abrirModal(
                        "SECCIÓN",
                        "Editar sección"
                    );

                    formularioSeccion(
                        this.dataset.nombre || "",
                        this.dataset.descripcion || ""
                    );
                }
            );
        }
    );


document
    .querySelectorAll(".agregar-subseccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    modoFormulario =
                        "crear_subseccion";

                    idFormulario =
                        null;

                    seccionPadreFormulario =
                        Number(
                            this.dataset.seccionId
                        );

                    abrirModal(
                        "SUBSECCIÓN",
                        `Nueva subsección · ${
                            this.dataset.seccionNombre
                            || ""
                        }`
                    );

                    formularioSubseccion();
                }
            );
        }
    );


document
    .querySelectorAll(".editar-subseccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    modoFormulario =
                        "editar_subseccion";

                    idFormulario =
                        Number(
                            this.dataset.id
                        );

                    abrirModal(
                        "SUBSECCIÓN",
                        "Editar subsección"
                    );

                    formularioSubseccion(
                        this.dataset.nombre || "",
                        this.dataset.descripcion || ""
                    );
                }
            );
        }
    );


document
    .querySelectorAll(".estado-seccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    cambiarEstado(
                        "secciones",
                        this.dataset.id,
                        this.dataset.accion
                    );
                }
            );
        }
    );


document
    .querySelectorAll(".estado-subseccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    cambiarEstado(
                        "subsecciones",
                        this.dataset.id,
                        this.dataset.accion
                    );
                }
            );
        }
    );


document
    .querySelectorAll(".eliminar-seccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    eliminarEntidad(
                        "secciones",
                        this.dataset.id,
                        `¿Eliminar definitivamente la sección "${this.dataset.nombre}"?\n\nSolo se eliminará si no contiene subsecciones ni materiales.`
                    );
                }
            );
        }
    );


document
    .querySelectorAll(".eliminar-subseccion")
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    eliminarEntidad(
                        "subsecciones",
                        this.dataset.id,
                        `¿Eliminar definitivamente la subsección "${this.dataset.nombre}"?\n\nSolo se eliminará si no contiene materiales.`
                    );
                }
            );
        }
    );


if (agregarMaterial) {
    agregarMaterial.addEventListener(
        "click",
        function () {
            modoFormulario =
                "crear_material";

            idFormulario =
                null;

            materialSeleccionado =
                null;

            abrirModal(
                "MATERIAL",
                "Agregar al inventario técnico"
            );

            formularioMaterial();
        }
    );
}


document
    .querySelectorAll(
        ".editar-material-inventario"
    )
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    modoFormulario =
                        "editar_material";

                    idFormulario =
                        Number(
                            this.dataset.id
                        );

                    abrirModal(
                        "MATERIAL",
                        "Editar control de stock"
                    );

                    formularioMaterial(
                        {
                            seccion_id:
                                this.dataset.seccionId,

                            subseccion_id:
                                this.dataset.subseccionId,

                            stock_minimo:
                                this.dataset.stockMinimo,

                            stock_objetivo:
                                this.dataset.stockObjetivo,

                            observaciones:
                                this.dataset.observaciones
                                || ""
                        }
                    );
                }
            );
        }
    );


document
    .querySelectorAll(
        ".estado-material-inventario"
    )
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    cambiarEstado(
                        "materiales",
                        this.dataset.id,
                        this.dataset.accion
                    );
                }
            );
        }
    );


document
    .querySelectorAll(
        ".eliminar-material-inventario"
    )
    .forEach(
        boton => {
            boton.addEventListener(
                "click",
                function () {
                    eliminarEntidad(
                        "materiales",
                        this.dataset.id,
                        `¿Eliminar definitivamente "${this.dataset.descripcion}" del inventario técnico?`
                    );
                }
            );
        }
    );


function aplicarFiltros() {
    const texto =
        buscarInventario
            ? buscarInventario.value
                .trim()
                .toLowerCase()
            : "";

    const estado =
        filtrarEstadoStock
            ? filtrarEstadoStock.value
            : "";

    document
        .querySelectorAll(
            "#tablaInventario tr[data-busqueda]"
        )
        .forEach(
            fila => {
                const busqueda =
                    String(
                        fila.dataset.busqueda || ""
                    ).toLowerCase();

                const estadoFila =
                    String(
                        fila.dataset.estado || ""
                    );

                const coincideTexto =
                    !texto ||
                    busqueda.includes(texto);

                const coincideEstado =
                    !estado ||
                    estadoFila === estado;

                fila.style.display =
                    coincideTexto &&
                    coincideEstado
                        ? ""
                        : "none";
            }
        );
}


if (buscarInventario) {
    buscarInventario.addEventListener(
        "input",
        aplicarFiltros
    );
}


if (filtrarEstadoStock) {
    filtrarEstadoStock.addEventListener(
        "change",
        aplicarFiltros
    );
}


if (guardarFormulario) {
    guardarFormulario.addEventListener(
        "click",
        guardar
    );
}


if (cerrarFormulario) {
    cerrarFormulario.addEventListener(
        "click",
        cerrarModal
    );
}


if (cancelarFormulario) {
    cancelarFormulario.addEventListener(
        "click",
        cerrarModal
    );
}


if (fondoFormulario) {
    fondoFormulario.addEventListener(
        "click",
        cerrarModal
    );
}


document.addEventListener(
    "keydown",
    function (evento) {
        if (
            evento.key === "Escape" &&
            modalFormulario &&
            !modalFormulario.classList.contains(
                "oculto"
            )
        ) {
            cerrarModal();
        }
    }
);