"use strict";

const buscarMaterial = document.getElementById("buscarMaterial");
const resultadosBusqueda = document.getElementById("resultadosBusqueda");
const bloqueCrearManual = document.getElementById("bloqueCrearManual");
const crearMaterialManual = document.getElementById("crearMaterialManual");
const totalResultados = document.getElementById("totalResultados");

const modalRepuesto = document.getElementById("modalRepuesto");
const cerrarModal = document.getElementById("cerrarModal");
const cerrarFondo = document.getElementById("cerrarFondo");
const cancelarModal = document.getElementById("cancelarModal");
const formAgregarMaterial = document.getElementById("formAgregarMaterial");

const codigoSap = document.getElementById("codigoSap");
const descripcionMaterial = document.getElementById("descripcionMaterial");
const unidadMedida = document.getElementById("unidadMedida");
const existencia = document.getElementById("existencia");
const origenDato = document.getElementById("origenDato");

const descripcionVisual = document.getElementById("descripcionVisual");
const codigoVisual = document.getElementById("codigoVisual");
const existenciaVisual = document.getElementById("existenciaVisual");
const unidadVisual = document.getElementById("unidadVisual");

const camposManual = document.getElementById("camposManual");
const codigoSapManual = document.getElementById("codigoSapManual");
const descripcionManual = document.getElementById("descripcionManual");
const unidadManual = document.getElementById("unidadManual");

const origen = document.getElementById("origen");
const consumoPromedio = document.getElementById("consumoPromedio");
const tiempoEntrega = document.getElementById("tiempoEntrega");
const cantidadPedir = document.getElementById("cantidadPedir");
const valorUnitario = document.getElementById("valorUnitario");
const observaciones = document.getElementById("observaciones");

const alertaPendiente = document.getElementById("alertaPendiente");

const tablaPedido = document.getElementById("tablaPedido");
const contadorMateriales = document.getElementById("contadorMateriales");
const cantidadItems = document.getElementById("cantidadItems");
const valorAproximado = document.getElementById("valorAproximado");

const generarPedido = document.getElementById("generarPedido");
const limpiarPedido = document.getElementById("limpiarPedido");
const resultadoGuardado = document.getElementById("resultadoGuardado");

let pedido = [];
let temporizadorBusqueda = null;
let modoManual = false;
let ultimoResultado = [];
let materialParaSeccion = null;
let catalogoSecciones = [];
let catalogoSubsecciones = [];

function escaparHtml(texto) {
    return String(texto ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatearNumero(valor) {
    return Number(valor || 0).toLocaleString(
        "es-CO",
        {
            maximumFractionDigits: 2
        }
    );
}

function formatearMoneda(valor) {
    return Number(valor || 0).toLocaleString(
        "es-CO",
        {
            style: "currency",
            currency: "COP",
            maximumFractionDigits: 0
        }
    );
}

function obtenerValorUnitario(material) {
    let valor = Number(
        material.valor_unitario || 0
    );

    if (
        valor <= 0 &&
        Number(material.existencia || 0) > 0 &&
        Number(material.valor_inventario || 0) > 0
    ) {
        valor =
            Number(material.valor_inventario) /
            Number(material.existencia);
    }

    return valor;
}

async function buscar() {
    const texto = buscarMaterial.value.trim();

    bloqueCrearManual.classList.add("oculto");

    if (texto.length < 2) {
        resultadosBusqueda.innerHTML = "";

        if (totalResultados) {
            totalResultados.textContent =
                "Escriba al menos 2 caracteres";
        }

        return;
    }

    resultadosBusqueda.innerHTML = `
        <div class="resultado-cargando">
            Buscando materiales en SAP...
        </div>
    `;

    if (totalResultados) {
        totalResultados.textContent =
            "Buscando...";
    }

    try {
        const respuesta = await fetch(
            `/repuestos/api/materiales?q=${encodeURIComponent(texto)}`
        );

        if (!respuesta.ok) {
            throw new Error(
                "No fue posible consultar la base SAP."
            );
        }

        const materiales = await respuesta.json();

        ultimoResultado =
            Array.isArray(materiales)
                ? materiales
                : [];

        pintarResultados(
            ultimoResultado
        );

    } catch (error) {
        resultadosBusqueda.innerHTML = `
            <div class="resultado-error">
                ${escaparHtml(error.message)}
            </div>
        `;

        if (totalResultados) {
            totalResultados.textContent =
                "Error";
        }
    }
}

function pintarResultados(materiales) {
    if (totalResultados) {
        totalResultados.textContent =
            materiales.length === 1
                ? "1 resultado"
                : `${materiales.length} resultados`;
    }

    if (!materiales.length) {
        resultadosBusqueda.innerHTML = `
            <div class="sin-resultados">
                No se encontraron materiales con ese código o descripción.
            </div>
        `;

        bloqueCrearManual.classList.remove("oculto");

        return;
    }

    bloqueCrearManual.classList.add("oculto");

    resultadosBusqueda.innerHTML =
        materiales
            .map(
                (material, indice) => {
                    const valor =
                        obtenerValorUnitario(material);

                    return `
                        <div
                            class="resultado-material"
                            data-indice="${indice}"
                        >

                            <div class="resultado-codigo">
                                ${escaparHtml(
                                    material.codigo_sap || "SIN CÓDIGO"
                                )}
                            </div>

                            <div class="resultado-descripcion">
                                ${escaparHtml(
                                    material.descripcion || ""
                                )}
                            </div>

                            <div class="resultado-dato">
                                ${escaparHtml(
                                    material.unidad_medida || "-"
                                )}
                            </div>

                            <div class="resultado-dato">
                                ${formatearNumero(
                                    material.existencia
                                )}
                            </div>

                            <div class="resultado-dato">
                                ${formatearMoneda(valor)}
                            </div>

                            <div class="acciones-resultado">

                                <button
                                    type="button"
                                    class="accion-pedido"
                                    data-indice="${indice}"
                                >
                                    🛒 Agregar a pedido
                                </button>

                                <button
                                    type="button"
                                    class="accion-seccion"
                                    data-indice="${indice}"
                                >
                                    📁 Agregar a sección
                                </button>

                            </div>

                        </div>
                    `;
                }
            )
            .join("");

    document
        .querySelectorAll(".accion-pedido")
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const material =
                            ultimoResultado[indice];

                        if (material) {
                            abrirMaterial(
                                material
                            );
                        }
                    }
                );
            }
        );

    document
        .querySelectorAll(".accion-seccion")
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const material =
                            ultimoResultado[indice];

                        if (material) {
                            abrirAgregarSeccion(
                                material
                            );
                        }
                    }
                );
            }
        );
}

async function abrirMaterial(material) {
    modoManual = false;

    camposManual.classList.add("oculto");

    origenDato.value = "SAP";

    codigoSap.value =
        material.codigo_sap || "";

    descripcionMaterial.value =
        material.descripcion || "";

    unidadMedida.value =
        material.unidad_medida || "";

    existencia.value =
        Number(material.existencia || 0);

    codigoVisual.textContent =
        material.codigo_sap || "SIN CÓDIGO";

    descripcionVisual.textContent =
        material.descripcion || "Sin descripción";

    unidadVisual.textContent =
        material.unidad_medida || "-";

    existenciaVisual.textContent =
        `${formatearNumero(
            material.existencia
        )} ${material.unidad_medida || ""}`;

    origen.value = "";
    consumoPromedio.value = "";
    tiempoEntrega.value = "";
    cantidadPedir.value = "";
    observaciones.value = "";

    valorUnitario.value =
        obtenerValorUnitario(
            material
        ).toFixed(2);

    alertaPendiente.classList.add("oculto");
    alertaPendiente.innerHTML = "";

    if (material.codigo_sap) {
        try {
            const respuesta = await fetch(
                `/repuestos/api/material/${encodeURIComponent(
                    material.codigo_sap
                )}/pendiente`
            );

            if (respuesta.ok) {
                const datos =
                    await respuesta.json();

                if (
                    datos.pendientes &&
                    datos.pendientes.length
                ) {
                    const totalPendiente =
                        datos.pendientes.reduce(
                            (acumulado, item) =>
                                acumulado +
                                Number(
                                    item.cantidad_pendiente || 0
                                ),
                            0
                        );

                    alertaPendiente.innerHTML = `
                        <strong>
                            ⚠ Este material ya tiene un pedido pendiente
                        </strong>

                        <span>
                            Pendiente por recibir:
                            ${formatearNumero(totalPendiente)}
                            ${escaparHtml(
                                material.unidad_medida || ""
                            )}
                        </span>
                    `;

                    alertaPendiente.classList.remove(
                        "oculto"
                    );
                }
            }

        } catch (error) {
            console.error(error);
        }
    }

    modalRepuesto.classList.remove("oculto");
}

function abrirManual() {
    modoManual = true;

    origenDato.value = "MANUAL";

    codigoSap.value = "";
    descripcionMaterial.value = "";
    unidadMedida.value = "";
    existencia.value = 0;

    codigoVisual.textContent =
        "CREAR";

    descripcionVisual.textContent =
        "Nuevo material";

    unidadVisual.textContent =
        "-";

    existenciaVisual.textContent =
        "0";

    codigoSapManual.value =
        "CREAR";

    descripcionManual.value =
        buscarMaterial.value.trim();

    unidadManual.value = "";

    camposManual.classList.remove(
        "oculto"
    );

    origen.value = "";
    consumoPromedio.value = "";
    tiempoEntrega.value = "";
    cantidadPedir.value = "";
    valorUnitario.value = "";
    observaciones.value = "";

    alertaPendiente.classList.add(
        "oculto"
    );

    modalRepuesto.classList.remove(
        "oculto"
    );
}

function cerrarVentana() {
    modalRepuesto.classList.add(
        "oculto"
    );
}

function agregarMaterialAlPedido(evento) {
    evento.preventDefault();

    if (modoManual) {
        const descripcion =
            descripcionManual.value.trim();

        const unidad =
            unidadManual.value.trim();

        if (!descripcion) {
            alert(
                "Ingrese la descripción del material."
            );

            return;
        }

        if (!unidad) {
            alert(
                "Ingrese la unidad de medida."
            );

            return;
        }

        codigoSap.value =
            codigoSapManual.value.trim() ||
            "CREAR";

        descripcionMaterial.value =
            descripcion;

        unidadMedida.value =
            unidad;

        existencia.value =
            0;
    }

    if (!formAgregarMaterial.reportValidity()) {
        return;
    }

    const codigo =
        codigoSap.value.trim();

    const existente =
        pedido.find(
            item =>
                item.codigo_sap === codigo &&
                codigo !== "CREAR"
        );

    if (existente) {
        alert(
            "Este material ya está agregado al pedido. Modifique la cantidad directamente en Pedido actual."
        );

        cerrarVentana();

        return;
    }

    pedido.push(
        {
            codigo_sap:
                codigo || "CREAR",

            descripcion:
                descripcionMaterial.value.trim(),

            origen:
                origen.value,

            unidad_medida:
                unidadMedida.value.trim(),

            existencia:
                Number(
                    existencia.value || 0
                ),

            consumo_promedio:
                Number(
                    consumoPromedio.value || 0
                ),

            tiempo_entrega:
                Number(
                    tiempoEntrega.value || 0
                ),

            cantidad_pedir:
                Number(
                    cantidadPedir.value || 0
                ),

            valor_unitario:
                Number(
                    valorUnitario.value || 0
                ),

            observaciones:
                observaciones.value.trim(),

            origen_dato:
                origenDato.value || "SAP"
        }
    );

    pintarPedido();

    cerrarVentana();
}

function actualizarResumenPedido() {
    const total =
        pedido.reduce(
            (acumulado, item) =>
                acumulado +
                (
                    Number(
                        item.cantidad_pedir || 0
                    ) *
                    Number(
                        item.valor_unitario || 0
                    )
                ),
            0
        );

    if (contadorMateriales) {
        contadorMateriales.textContent =
            `${pedido.length} ${
                pedido.length === 1
                    ? "material"
                    : "materiales"
            }`;
    }

    if (cantidadItems) {
        cantidadItems.textContent =
            pedido.length;
    }

    if (valorAproximado) {
        valorAproximado.textContent =
            formatearMoneda(total);
    }

    generarPedido.disabled =
        pedido.length === 0;
}

function pintarPedido() {
    if (!pedido.length) {
        tablaPedido.innerHTML = `
            <tr>
                <td
                    colspan="5"
                    class="tabla-vacia"
                >
                    Los repuestos agregados al pedido aparecerán aquí.
                </td>
            </tr>
        `;

        actualizarResumenPedido();

        return;
    }

    tablaPedido.innerHTML =
        pedido
            .map(
                (item, indice) => `
                    <tr>

                        <td>
                            ${escaparHtml(
                                item.codigo_sap || "CREAR"
                            )}
                        </td>

                        <td>
                            ${escaparHtml(
                                item.descripcion
                            )}
                        </td>

                        <td>
                            <input
                                type="number"
                                class="editar-cantidad"
                                data-indice="${indice}"
                                min="0.01"
                                step="any"
                                value="${Number(
                                    item.cantidad_pedir || 0
                                )}"
                            >
                        </td>

                        <td>
                            <input
                                type="number"
                                class="editar-valor"
                                data-indice="${indice}"
                                min="0"
                                step="any"
                                value="${Number(
                                    item.valor_unitario || 0
                                )}"
                            >
                        </td>

                        <td>

                            <div class="acciones-linea">

                                <button
                                    type="button"
                                    class="boton-editar"
                                    data-indice="${indice}"
                                    title="Editar información"
                                >
                                    ✎
                                </button>

                                <button
                                    type="button"
                                    class="boton-eliminar"
                                    data-indice="${indice}"
                                    title="Eliminar"
                                >
                                    ✕
                                </button>

                            </div>

                        </td>

                    </tr>
                `
            )
            .join("");

    document
        .querySelectorAll(".editar-cantidad")
        .forEach(
            campo => {
                campo.addEventListener(
                    "input",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const valor =
                            Number(this.value);

                        if (
                            Number.isFinite(valor) &&
                            valor > 0
                        ) {
                            pedido[
                                indice
                            ].cantidad_pedir =
                                valor;

                            actualizarResumenPedido();
                        }
                    }
                );
            }
        );

    document
        .querySelectorAll(".editar-valor")
        .forEach(
            campo => {
                campo.addEventListener(
                    "input",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const valor =
                            Number(this.value);

                        if (
                            Number.isFinite(valor) &&
                            valor >= 0
                        ) {
                            pedido[
                                indice
                            ].valor_unitario =
                                valor;

                            actualizarResumenPedido();
                        }
                    }
                );
            }
        );

    document
        .querySelectorAll(".boton-eliminar")
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        pedido.splice(
                            indice,
                            1
                        );

                        pintarPedido();
                    }
                );
            }
        );

    document
        .querySelectorAll(".boton-editar")
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    function () {
                        editarLineaPedido(
                            Number(
                                this.dataset.indice
                            )
                        );
                    }
                );
            }
        );

    actualizarResumenPedido();
}

function editarLineaPedido(indice) {
    const item =
        pedido[indice];

    if (!item) {
        return;
    }

    const consumo =
        prompt(
            "Consumo promedio mensual:",
            item.consumo_promedio
        );

    if (consumo === null) {
        return;
    }

    const tiempo =
        prompt(
            "Tiempo de entrega en meses:",
            item.tiempo_entrega
        );

    if (tiempo === null) {
        return;
    }

    const observacion =
        prompt(
            "Observaciones:",
            item.observaciones || ""
        );

    if (observacion === null) {
        return;
    }

    item.consumo_promedio =
        Number(consumo || 0);

    item.tiempo_entrega =
        Number(tiempo || 0);

    item.observaciones =
        observacion.trim();

    pintarPedido();
}

function crearModalSeccion() {
    if (
        document.getElementById(
            "modalAgregarSeccion"
        )
    ) {
        return;
    }

    const modal =
        document.createElement("div");

    modal.id =
        "modalAgregarSeccion";

    modal.className =
        "modal-repuesto oculto";

    modal.innerHTML = `
        <div
            class="fondo-modal-repuesto"
            id="fondoAgregarSeccion"
        ></div>

        <section class="contenido-modal-repuesto">

            <button
                type="button"
                class="cerrar-modal-repuesto"
                id="cerrarAgregarSeccion"
            >
                ×
            </button>

            <span class="eyebrow">
                INVENTARIO TÉCNICO
            </span>

            <h2>
                Agregar material a sección
            </h2>

            <section class="informacion-material">

                <div class="info-material-principal">

                    <span class="info-etiqueta">
                        MATERIAL
                    </span>

                    <strong id="seccionMaterialDescripcion">
                        -
                    </strong>

                </div>

                <div class="info-material-datos">

                    <div>
                        <span>Código SAP</span>
                        <strong id="seccionMaterialCodigo">-</strong>
                    </div>

                    <div>
                        <span>Existencia SAP</span>
                        <strong id="seccionMaterialStock">0</strong>
                    </div>

                    <div>
                        <span>Unidad</span>
                        <strong id="seccionMaterialUnidad">-</strong>
                    </div>

                </div>

            </section>

            <div class="rejilla-formulario">

                <div class="campo">

                    <label>
                        Sección
                    </label>

                    <select id="selectorSeccion">
                        <option value="">
                            Seleccione
                        </option>
                    </select>

                </div>

                <div class="campo">

                    <label>
                        Subsección
                    </label>

                    <select id="selectorSubseccion">
                        <option value="">
                            Sin subsección
                        </option>
                    </select>

                </div>

                <div class="campo">

                    <label>
                        Stock mínimo
                    </label>

                    <input
                        type="number"
                        id="stockMinimoSeccion"
                        min="0"
                        step="any"
                        value="0"
                    >

                </div>

                <div class="campo">

                    <label>
                        Stock objetivo
                    </label>

                    <input
                        type="number"
                        id="stockObjetivoSeccion"
                        min="0"
                        step="any"
                        value="0"
                    >

                </div>

                <div class="campo campo-ancho">

                    <label>
                        Observaciones
                    </label>

                    <textarea
                        id="observacionesSeccion"
                        rows="3"
                    ></textarea>

                </div>

            </div>

            <div
                id="errorAgregarSeccion"
                class="alerta-pedido oculto"
            ></div>

            <div class="acciones-modal">

                <button
                    type="button"
                    class="boton-secundario"
                    id="cancelarAgregarSeccion"
                >
                    Cancelar
                </button>

                <button
                    type="button"
                    class="boton-principal"
                    id="guardarAgregarSeccion"
                >
                    Agregar a sección
                </button>

            </div>

        </section>
    `;

    document.body.appendChild(
        modal
    );

    document
        .getElementById(
            "cerrarAgregarSeccion"
        )
        .addEventListener(
            "click",
            cerrarModalSeccion
        );

    document
        .getElementById(
            "cancelarAgregarSeccion"
        )
        .addEventListener(
            "click",
            cerrarModalSeccion
        );

    document
        .getElementById(
            "fondoAgregarSeccion"
        )
        .addEventListener(
            "click",
            cerrarModalSeccion
        );

    document
        .getElementById(
            "selectorSeccion"
        )
        .addEventListener(
            "change",
            actualizarSubsecciones
        );

    document
        .getElementById(
            "guardarAgregarSeccion"
        )
        .addEventListener(
            "click",
            guardarMaterialEnSeccion
        );
}

async function cargarCatalogoInventario() {
    const respuesta =
        await fetch(
            "/repuestos/api/inventario/catalogo"
        );

    if (!respuesta.ok) {
        throw new Error(
            "No fue posible consultar las secciones del inventario técnico."
        );
    }

    const datos =
        await respuesta.json();

    if (!datos.ok) {
        throw new Error(
            datos.error ||
            "No fue posible consultar el inventario."
        );
    }

    catalogoSecciones =
        datos.secciones || [];

    catalogoSubsecciones =
        datos.subsecciones || [];
}

async function abrirAgregarSeccion(material) {
    materialParaSeccion =
        material;

    crearModalSeccion();

    const modal =
        document.getElementById(
            "modalAgregarSeccion"
        );

    const error =
        document.getElementById(
            "errorAgregarSeccion"
        );

    error.classList.add(
        "oculto"
    );

    error.innerHTML = "";

    document.getElementById(
        "seccionMaterialDescripcion"
    ).textContent =
        material.descripcion || "-";

    document.getElementById(
        "seccionMaterialCodigo"
    ).textContent =
        material.codigo_sap || "CREAR";

    document.getElementById(
        "seccionMaterialStock"
    ).textContent =
        formatearNumero(
            material.existencia
        );

    document.getElementById(
        "seccionMaterialUnidad"
    ).textContent =
        material.unidad_medida || "-";

    try {
        await cargarCatalogoInventario();

        const selector =
            document.getElementById(
                "selectorSeccion"
            );

        selector.innerHTML = `
            <option value="">
                Seleccione
            </option>

            ${catalogoSecciones
                .filter(
                    seccion =>
                        Number(
                            seccion.activo
                        ) === 1
                )
                .map(
                    seccion => `
                        <option value="${seccion.id}">
                            ${escaparHtml(
                                seccion.nombre
                            )}
                        </option>
                    `
                )
                .join("")}
        `;

        document.getElementById(
            "selectorSubseccion"
        ).innerHTML = `
            <option value="">
                Sin subsección
            </option>
        `;

        modal.classList.remove(
            "oculto"
        );

    } catch (errorCarga) {
        alert(
            errorCarga.message
        );
    }
}

function actualizarSubsecciones() {
    const seccionId =
        Number(
            document.getElementById(
                "selectorSeccion"
            ).value
        );

    const selector =
        document.getElementById(
            "selectorSubseccion"
        );

    selector.innerHTML = `
        <option value="">
            Sin subsección
        </option>

        ${catalogoSubsecciones
            .filter(
                sub =>
                    Number(sub.activo) === 1 &&
                    Number(sub.seccion_id) === seccionId
            )
            .map(
                sub => `
                    <option value="${sub.id}">
                        ${escaparHtml(
                            sub.nombre
                        )}
                    </option>
                `
            )
            .join("")}
    `;
}

function cerrarModalSeccion() {
    const modal =
        document.getElementById(
            "modalAgregarSeccion"
        );

    if (modal) {
        modal.classList.add(
            "oculto"
        );
    }

    materialParaSeccion =
        null;
}

async function guardarMaterialEnSeccion() {
    if (!materialParaSeccion) {
        return;
    }

    const error =
        document.getElementById(
            "errorAgregarSeccion"
        );

    error.classList.add(
        "oculto"
    );

    const seccionId =
        document.getElementById(
            "selectorSeccion"
        ).value;

    const subseccionId =
        document.getElementById(
            "selectorSubseccion"
        ).value;

    const stockMinimo =
        Number(
            document.getElementById(
                "stockMinimoSeccion"
            ).value || 0
        );

    const stockObjetivo =
        Number(
            document.getElementById(
                "stockObjetivoSeccion"
            ).value || 0
        );

    if (!seccionId) {
        error.innerHTML =
            "Seleccione una sección.";

        error.classList.remove(
            "oculto"
        );

        return;
    }

    if (stockMinimo < 0) {
        error.innerHTML =
            "El stock mínimo no puede ser negativo.";

        error.classList.remove(
            "oculto"
        );

        return;
    }

    if (
        stockObjetivo <
        stockMinimo
    ) {
        error.innerHTML =
            "El stock objetivo debe ser igual o mayor al stock mínimo.";

        error.classList.remove(
            "oculto"
        );

        return;
    }

    const boton =
        document.getElementById(
            "guardarAgregarSeccion"
        );

    boton.disabled = true;
    boton.textContent =
        "Guardando...";

    try {
        const respuesta =
            await fetch(
                "/repuestos/api/inventario/materiales",
                {
                    method:
                        "POST",

                    headers:
                        {
                            "Content-Type":
                                "application/json"
                        },

                    body:
                        JSON.stringify(
                            {
                                codigo_sap:
                                    materialParaSeccion.codigo_sap,

                                descripcion:
                                    materialParaSeccion.descripcion,

                                unidad_medida:
                                    materialParaSeccion.unidad_medida,

                                seccion_id:
                                    seccionId,

                                subseccion_id:
                                    subseccionId,

                                stock_minimo:
                                    stockMinimo,

                                stock_objetivo:
                                    stockObjetivo,

                                observaciones:
                                    document.getElementById(
                                        "observacionesSeccion"
                                    ).value.trim(),

                                origen_dato:
                                    "SAP"
                            }
                        )
                }
            );

        const datos =
            await respuesta.json();

        if (
            !respuesta.ok ||
            !datos.ok
        ) {
            throw new Error(
                datos.error ||
                "No fue posible agregar el material a la sección."
            );
        }

        cerrarModalSeccion();

        alert(
            "Material agregado correctamente al Inventario técnico."
        );

    } catch (errorGuardar) {
        error.innerHTML =
            escaparHtml(
                errorGuardar.message
            );

        error.classList.remove(
            "oculto"
        );

    } finally {
        boton.disabled = false;
        boton.textContent =
            "Agregar a sección";
    }
}

async function generarSolicitud() {
    if (!pedido.length) {
        return;
    }

    generarPedido.disabled =
        true;

    generarPedido.textContent =
        "Generando...";

    resultadoGuardado.classList.add(
        "oculto"
    );

    resultadoGuardado.innerHTML =
        "";

    try {
        const respuesta =
            await fetch(
                "/repuestos/api/solicitudes",
                {
                    method:
                        "POST",

                    headers:
                        {
                            "Content-Type":
                                "application/json"
                        },

                    body:
                        JSON.stringify(
                            {
                                items:
                                    pedido
                            }
                        )
                }
            );

        const datos =
            await respuesta.json();

        if (
            !respuesta.ok ||
            !datos.ok
        ) {
            throw new Error(
                datos.error ||
                "No fue posible generar el pedido."
            );
        }

        const solicitud =
            datos.solicitud;

        resultadoGuardado.innerHTML = `
            <strong>
                ✓ Pedido generado correctamente
            </strong>

            <div style="margin-top:5px;">
                ${escaparHtml(
                    solicitud.codigo_solicitud || ""
                )}
            </div>
        `;

        resultadoGuardado.classList.remove(
            "oculto"
        );

        if (
            solicitud.archivo_url
        ) {
            const enlace =
                document.createElement(
                    "a"
                );

            enlace.href =
                solicitud.archivo_url;

            enlace.style.display =
                "none";

            document.body.appendChild(
                enlace
            );

            enlace.click();

            enlace.remove();
        }

        pedido = [];

        pintarPedido();

    } catch (error) {
        resultadoGuardado.innerHTML = `
            <strong>
                ✕ No fue posible generar el pedido
            </strong>

            <div style="margin-top:5px;">
                ${escaparHtml(
                    error.message
                )}
            </div>
        `;

        resultadoGuardado.classList.remove(
            "oculto"
        );

    } finally {
        generarPedido.textContent =
            "▣ Generar pedido";

        generarPedido.disabled =
            pedido.length === 0;
    }
}

if (buscarMaterial) {
    buscarMaterial.addEventListener(
        "input",
        function () {
            clearTimeout(
                temporizadorBusqueda
            );

            temporizadorBusqueda =
                setTimeout(
                    buscar,
                    250
                );
        }
    );
}

if (crearMaterialManual) {
    crearMaterialManual.addEventListener(
        "click",
        abrirManual
    );
}

if (cerrarModal) {
    cerrarModal.addEventListener(
        "click",
        cerrarVentana
    );
}

if (cerrarFondo) {
    cerrarFondo.addEventListener(
        "click",
        cerrarVentana
    );
}

if (cancelarModal) {
    cancelarModal.addEventListener(
        "click",
        cerrarVentana
    );
}

if (formAgregarMaterial) {
    formAgregarMaterial.addEventListener(
        "submit",
        agregarMaterialAlPedido
    );
}

if (limpiarPedido) {
    limpiarPedido.addEventListener(
        "click",
        function () {
            if (!pedido.length) {
                return;
            }

            if (
                !confirm(
                    "¿Desea limpiar todo el pedido actual?"
                )
            ) {
                return;
            }

            pedido = [];

            pintarPedido();
        }
    );
}

if (generarPedido) {
    generarPedido.addEventListener(
        "click",
        generarSolicitud
    );
}

pintarPedido();