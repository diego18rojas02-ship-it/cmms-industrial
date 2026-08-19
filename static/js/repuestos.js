"use strict";

const buscarMaterial = document.getElementById("buscarMaterial");
const resultadosBusqueda = document.getElementById("resultadosBusqueda");
const bloqueCrearManual = document.getElementById("bloqueCrearManual");
const crearMaterialManual = document.getElementById("crearMaterialManual");

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
const valorAproximado = document.getElementById("valorAproximado");
const generarPedido = document.getElementById("generarPedido");
const limpiarPedido = document.getElementById("limpiarPedido");
const resultadoGuardado = document.getElementById("resultadoGuardado");

let pedido = [];
let temporizadorBusqueda = null;
let modoManual = false;

function escaparHtml(texto) {
    return String(texto ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
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

async function buscar() {
    const texto = buscarMaterial.value.trim();

    bloqueCrearManual.classList.add("oculto");

    if (texto.length < 2) {
        resultadosBusqueda.innerHTML = "";
        return;
    }

    resultadosBusqueda.innerHTML = `
        <div class="resultado-cargando">
            Buscando...
        </div>
    `;

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

        pintarResultados(materiales);

    } catch (error) {
        resultadosBusqueda.innerHTML = `
            <div class="resultado-error">
                ${escaparHtml(error.message)}
            </div>
        `;
    }
}

function pintarResultados(materiales) {
    if (!materiales.length) {
        resultadosBusqueda.innerHTML = `
            <div class="sin-resultados">
                <strong>
                    No se encontraron materiales.
                </strong>

                <span>
                    Puede crear el ítem manualmente.
                </span>
            </div>
        `;

        bloqueCrearManual.classList.remove("oculto");

        return;
    }

    bloqueCrearManual.classList.add("oculto");

    resultadosBusqueda.innerHTML =
        materiales
            .map(
                material => `
                    <button
                        type="button"
                        class="resultado-material"
                        data-id="${material.id}"
                    >
                        <div>
                            <strong>
                                ${escaparHtml(
                                    material.codigo_sap || "SIN CÓDIGO"
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
                                    material.unidad_medida || "-"
                                )}
                            </span>

                            <strong>
                                ${formatearNumero(
                                    material.existencia
                                )}
                            </strong>

                            <span>
                                ${formatearMoneda(
                                    material.valor_unitario
                                )}
                            </span>
                        </div>
                    </button>
                `
            )
            .join("");

    document
        .querySelectorAll(".resultado-material")
        .forEach(
            boton => {
                boton.addEventListener(
                    "click",
                    function () {
                        const id = Number(
                            this.dataset.id
                        );

                        const material = materiales.find(
                            item =>
                                Number(item.id) === id
                        );

                        if (material) {
                            abrirMaterial(material);
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

    let valorSugerido =
        Number(
            material.valor_unitario || 0
        );

    if (
        !valorSugerido &&
        Number(material.existencia || 0) > 0
    ) {
        valorSugerido =
            Number(
                material.valor_inventario || 0
            ) /
            Number(
                material.existencia || 1
            );
    }

    valorUnitario.value =
        valorSugerido.toFixed(2);

    alertaPendiente.classList.add("oculto");
    alertaPendiente.innerHTML = "";

    if (material.codigo_sap) {
        try {
            const respuesta = await fetch(
                `/repuestos/api/material/${encodeURIComponent(
                    material.codigo_sap
                )}/pendiente`
            );

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
                        ⚠ Material con pedido pendiente
                    </strong>

                    <span>
                        Tiene
                        ${formatearNumero(totalPendiente)}
                        ${escaparHtml(
                            material.unidad_medida || ""
                        )}
                        pendientes por recibir.
                    </span>

                    <span>
                        Verifique antes de generar una nueva solicitud.
                    </span>
                `;

                alertaPendiente.classList.remove(
                    "oculto"
                );
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

    codigoVisual.textContent = "CREAR";
    descripcionVisual.textContent = "Ítem manual";
    unidadVisual.textContent = "-";
    existenciaVisual.textContent = "0";

    codigoSapManual.value = "CREAR";

    descripcionManual.value =
        buscarMaterial.value.trim();

    unidadManual.value = "";

    camposManual.classList.remove("oculto");

    origen.value = "";
    consumoPromedio.value = "";
    tiempoEntrega.value = "";
    cantidadPedir.value = "";
    valorUnitario.value = "";
    observaciones.value = "";

    alertaPendiente.classList.add("oculto");
    alertaPendiente.innerHTML = "";

    modalRepuesto.classList.remove("oculto");

    setTimeout(
        () => descripcionManual.focus(),
        80
    );
}

function cerrarVentana() {
    modalRepuesto.classList.add("oculto");
}

function agregarMaterialAlPedido(evento) {
    evento.preventDefault();

    if (modoManual) {
        const descripcion =
            descripcionManual.value.trim();

        const unidad =
            unidadManual.value.trim();

        if (!descripcion) {
            descripcionManual.setCustomValidity(
                "Ingrese la descripción."
            );

            descripcionManual.reportValidity();

            return;
        }

        descripcionManual.setCustomValidity("");

        if (!unidad) {
            unidadManual.setCustomValidity(
                "Ingrese la unidad de medida."
            );

            unidadManual.reportValidity();

            return;
        }

        unidadManual.setCustomValidity("");

        codigoSap.value =
            codigoSapManual.value.trim() ||
            "CREAR";

        descripcionMaterial.value =
            descripcion;

        unidadMedida.value =
            unidad;

        existencia.value = 0;
    }

    if (!formAgregarMaterial.reportValidity()) {
        return;
    }

    const codigo =
        codigoSap.value.trim();

    if (
        codigo &&
        codigo !== "CREAR" &&
        pedido.some(
            item =>
                item.codigo_sap === codigo
        )
    ) {
        alert(
            "Este código SAP ya está agregado al pedido actual. Puede modificar la cantidad directamente en la tabla."
        );

        return;
    }

    pedido.push(
        {
            codigo_sap:
                codigo,

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
    contadorMateriales.textContent =
        `${pedido.length} ${
            pedido.length === 1
                ? "material"
                : "materiales"
        }`;

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

    valorAproximado.textContent =
        formatearMoneda(total);

    generarPedido.disabled =
        pedido.length === 0;
}

function pintarPedido() {
    if (!pedido.length) {
        tablaPedido.innerHTML = `
            <tr>
                <td
                    colspan="7"
                    class="tabla-vacia"
                >
                    No hay materiales agregados.
                </td>
            </tr>
        `;

        contadorMateriales.textContent =
            "0 materiales";

        valorAproximado.textContent =
            "$0";

        generarPedido.disabled =
            true;

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
                            ${escaparHtml(
                                item.unidad_medida
                            )}
                        </td>

                        <td>
                            ${formatearNumero(
                                item.existencia
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
                                    title="Editar"
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
                            !Number.isFinite(valor) ||
                            valor <= 0
                        ) {
                            return;
                        }

                        pedido[indice].cantidad_pedir =
                            valor;

                        actualizarResumenPedido();
                    }
                );

                campo.addEventListener(
                    "change",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const valor =
                            Number(this.value);

                        if (
                            !Number.isFinite(valor) ||
                            valor <= 0
                        ) {
                            this.value =
                                pedido[indice]
                                    .cantidad_pedir;

                            return;
                        }

                        pedido[indice].cantidad_pedir =
                            valor;

                        actualizarResumenPedido();
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
                            !Number.isFinite(valor) ||
                            valor < 0
                        ) {
                            return;
                        }

                        pedido[indice].valor_unitario =
                            valor;

                        actualizarResumenPedido();
                    }
                );

                campo.addEventListener(
                    "change",
                    function () {
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        const valor =
                            Number(this.value);

                        if (
                            !Number.isFinite(valor) ||
                            valor < 0
                        ) {
                            this.value =
                                pedido[indice]
                                    .valor_unitario;

                            return;
                        }

                        pedido[indice].valor_unitario =
                            valor;

                        actualizarResumenPedido();
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
                        const indice =
                            Number(
                                this.dataset.indice
                            );

                        editarMaterialPedido(indice);
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

    actualizarResumenPedido();
}

function editarMaterialPedido(indice) {
    const item =
        pedido[indice];

    if (!item) {
        return;
    }

    const nuevaCantidad =
        prompt(
            `Cantidad a pedir para:\n${item.descripcion}`,
            item.cantidad_pedir
        );

    if (nuevaCantidad === null) {
        return;
    }

    const cantidad =
        Number(nuevaCantidad);

    if (
        !Number.isFinite(cantidad) ||
        cantidad <= 0
    ) {
        alert(
            "La cantidad debe ser mayor que cero."
        );

        return;
    }

    const nuevoConsumo =
        prompt(
            "Consumo promedio mensual:",
            item.consumo_promedio
        );

    if (nuevoConsumo === null) {
        return;
    }

    const consumo =
        Number(nuevoConsumo);

    if (
        !Number.isFinite(consumo) ||
        consumo < 0
    ) {
        alert(
            "El consumo promedio no es válido."
        );

        return;
    }

    const nuevoTiempo =
        prompt(
            "Tiempo de entrega:",
            item.tiempo_entrega
        );

    if (nuevoTiempo === null) {
        return;
    }

    const tiempo =
        Number(nuevoTiempo);

    if (
        !Number.isFinite(tiempo) ||
        tiempo < 0
    ) {
        alert(
            "El tiempo de entrega no es válido."
        );

        return;
    }

    const nuevoValor =
        prompt(
            "Valor unitario:",
            item.valor_unitario
        );

    if (nuevoValor === null) {
        return;
    }

    const valor =
        Number(nuevoValor);

    if (
        !Number.isFinite(valor) ||
        valor < 0
    ) {
        alert(
            "El valor unitario no es válido."
        );

        return;
    }

    const nuevaObservacion =
        prompt(
            "Observaciones:",
            item.observaciones || ""
        );

    if (nuevaObservacion === null) {
        return;
    }

    item.cantidad_pedir =
        cantidad;

    item.consumo_promedio =
        consumo;

    item.tiempo_entrega =
        tiempo;

    item.valor_unitario =
        valor;

    item.observaciones =
        nuevaObservacion.trim();

    pintarPedido();
}

async function generarSolicitud() {
    if (!pedido.length) {
        return;
    }

    const cantidadesInvalidas =
        pedido.some(
            item =>
                !Number.isFinite(
                    Number(item.cantidad_pedir)
                ) ||
                Number(item.cantidad_pedir) <= 0
        );

    if (cantidadesInvalidas) {
        alert(
            "Revise las cantidades a pedir antes de generar el pedido."
        );

        return;
    }

    generarPedido.disabled =
        true;

    generarPedido.textContent =
        "Generando...";

    resultadoGuardado.classList.add(
        "oculto"
    );

    resultadoGuardado.innerHTML = "";

    try {
        const respuesta = await fetch(
            "/repuestos/api/solicitudes",
            {
                method: "POST",

                headers: {
                    "Content-Type":
                        "application/json"
                },

                body: JSON.stringify(
                    {
                        items: pedido
                    }
                )
            }
        );

        let datos;

        try {
            datos =
                await respuesta.json();

        } catch (error) {
            throw new Error(
                "El servidor no devolvió una respuesta válida."
            );
        }

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

            <span>
                ${escaparHtml(
                    solicitud.codigo_solicitud
                )}
                · Semana
                ${solicitud.semana}
                ·
                ${solicitud.cantidad_items}
                ítems
                ·
                ${formatearMoneda(
                    solicitud.valor_estimado
                )}
            </span>
        `;

        resultadoGuardado.classList.remove(
            "oculto"
        );

        if (solicitud.archivo_url) {
            const enlace =
                document.createElement("a");

            enlace.href =
                solicitud.archivo_url;

            if (solicitud.archivo) {
                enlace.download =
                    solicitud.archivo;
            }

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

        buscarMaterial.value = "";

        resultadosBusqueda.innerHTML = "";

        bloqueCrearManual.classList.add(
            "oculto"
        );

    } catch (error) {
        resultadoGuardado.innerHTML = `
            <strong>
                ✕ No fue posible generar el pedido
            </strong>

            <span>
                ${escaparHtml(
                    error.message
                )}
            </span>
        `;

        resultadoGuardado.classList.remove(
            "oculto"
        );

    } finally {
        generarPedido.textContent =
            "Generar pedido";

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
                    280
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

            const confirmar =
                confirm(
                    "¿Desea eliminar todos los materiales del pedido?"
                );

            if (!confirmar) {
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