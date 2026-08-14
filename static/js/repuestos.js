"use strict";

const buscarMaterial =
    document.getElementById(
        "buscarMaterial"
    );

const resultadosBusqueda =
    document.getElementById(
        "resultadosBusqueda"
    );

const modalRepuesto =
    document.getElementById(
        "modalRepuesto"
    );

const cerrarModal =
    document.getElementById(
        "cerrarModal"
    );

const cerrarFondo =
    document.getElementById(
        "cerrarFondo"
    );

const cancelarModal =
    document.getElementById(
        "cancelarModal"
    );

const formAgregarMaterial =
    document.getElementById(
        "formAgregarMaterial"
    );

const codigoSap =
    document.getElementById(
        "codigoSap"
    );

const descripcionMaterial =
    document.getElementById(
        "descripcionMaterial"
    );

const unidadMedida =
    document.getElementById(
        "unidadMedida"
    );

const existencia =
    document.getElementById(
        "existencia"
    );

const consumoPromedio =
    document.getElementById(
        "consumoPromedio"
    );

const tiempoEntrega =
    document.getElementById(
        "tiempoEntrega"
    );

const cantidadPedir =
    document.getElementById(
        "cantidadPedir"
    );

const valorUnitario =
    document.getElementById(
        "valorUnitario"
    );

const observaciones =
    document.getElementById(
        "observaciones"
    );

const origen =
    document.getElementById(
        "origen"
    );

const alertaPendiente =
    document.getElementById(
        "alertaPendiente"
    );

const tablaPedido =
    document.getElementById(
        "tablaPedido"
    );

const contadorMateriales =
    document.getElementById(
        "contadorMateriales"
    );

const generarPedido =
    document.getElementById(
        "generarPedido"
    );

const limpiarPedido =
    document.getElementById(
        "limpiarPedido"
    );

const descripcionVisual =
    document.getElementById(
        "descripcionVisual"
    );

const codigoVisual =
    document.getElementById(
        "codigoVisual"
    );

const existenciaVisual =
    document.getElementById(
        "existenciaVisual"
    );

const unidadVisual =
    document.getElementById(
        "unidadVisual"
    );

let pedido = [];

let temporizadorBusqueda = null;


function escaparHtml(texto) {
    return String(
        texto ?? ""
    )
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}


function formatearNumero(valor) {
    return Number(
        valor || 0
    ).toLocaleString(
        "es-CO"
    );
}


async function buscar() {
    const texto =
        buscarMaterial.value.trim();

    if (
        texto.length < 2
    ) {
        resultadosBusqueda.innerHTML =
            "";

        return;
    }

    resultadosBusqueda.innerHTML =
        `
        <div class="resultado-cargando">
            Buscando...
        </div>
        `;

    try {
        const respuesta =
            await fetch(
                `/repuestos/api/materiales?q=${encodeURIComponent(texto)}`
            );

        const materiales =
            await respuesta.json();

        pintarResultados(
            materiales
        );

    } catch (error) {
        resultadosBusqueda.innerHTML =
            `
            <div class="resultado-error">
                No fue posible consultar materiales.
            </div>
            `;
    }
}


function pintarResultados(
    materiales
) {
    if (
        !materiales.length
    ) {
        resultadosBusqueda.innerHTML =
            `
            <div class="sin-resultados">

                <strong>
                    No se encontraron materiales.
                </strong>

                <span>
                    Más adelante habilitaremos la creación manual.
                </span>

            </div>
            `;

        return;
    }

    resultadosBusqueda.innerHTML =
        materiales
            .map(
                material => `
                    <button
                        type="button"
                        class="resultado-material"
                        data-material='${escaparHtml(
                            JSON.stringify(
                                material
                            )
                        )}'
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
                                ${formatearNumero(
                                    material.existencia
                                )}
                            </strong>

                        </div>

                    </button>
                `
            )
            .join("");

    document
        .querySelectorAll(
            ".resultado-material"
        )
        .forEach(
            boton => {

                boton.addEventListener(
                    "click",
                    function () {

                        const material =
                            JSON.parse(
                                this.dataset.material
                            );

                        abrirMaterial(
                            material
                        );

                    }
                );

            }
        );
}


async function abrirMaterial(
    material
) {
    codigoSap.value =
        material.codigo_sap
        || "";

    descripcionMaterial.value =
        material.descripcion
        || "";

    unidadMedida.value =
        material.unidad_medida
        || "";

    existencia.value =
        Number(
            material.existencia
            || 0
        );

        codigoVisual.textContent =
    material.codigo_sap
    || "CREAR";

descripcionVisual.textContent =
    material.descripcion
    || "Sin descripción";

unidadVisual.textContent =
    material.unidad_medida
    || "-";

existenciaVisual.textContent =
    `${formatearNumero(
        material.existencia
        || 0
    )} ${
        material.unidad_medida
        || ""
    }`;

    origen.value =
        "";

    consumoPromedio.value =
        "";

    tiempoEntrega.value =
        "";

    cantidadPedir.value =
        "";

    observaciones.value =
        "";

    let valorSugerido =
        Number(
            material.valor_unitario
            || 0
        );

    if (
        !valorSugerido
        && Number(
            material.existencia
        ) > 0
    ) {
        valorSugerido =
            Number(
                material.valor_inventario
                || 0
            )
            /
            Number(
                material.existencia
            );
    }

    valorUnitario.value =
        valorSugerido.toFixed(
            2
        );

    alertaPendiente.classList.add(
        "oculto"
    );

    alertaPendiente.innerHTML =
        "";

    if (
        material.codigo_sap
    ) {
        try {
            const respuesta =
                await fetch(
                    `/repuestos/api/material/${encodeURIComponent(material.codigo_sap)}/pendiente`
                );

            const datos =
                await respuesta.json();

            if (
                datos.pendiente
            ) {
                const pendiente =
                    datos.pendiente;

                alertaPendiente.innerHTML =
                    `
                    <strong>
                        ⚠ Pedido pendiente
                    </strong>

                    <span>
                        Solicitud:
                        ${escaparHtml(
                            pendiente.codigo_solicitud
                        )}
                    </span>

                    <span>
                        Cantidad pendiente:
                        ${formatearNumero(
                            pendiente.cantidad_pendiente
                        )}
                    </span>
                    `;

                alertaPendiente.classList.remove(
                    "oculto"
                );
            }

        } catch (error) {
            console.error(
                error
            );
        }
    }

    modalRepuesto.classList.remove(
        "oculto"
    );
}


function cerrarVentana() {
    modalRepuesto.classList.add(
        "oculto"
    );
}


function agregarMaterialAlPedido(
    evento
) {
    evento.preventDefault();

    if (
        !formAgregarMaterial.reportValidity()
    ) {
        return;
    }

    const codigo =
        codigoSap.value.trim();

    const materialExistente =
        pedido.find(
            item =>
                item.codigo_sap ===
                codigo
        );

    if (
        materialExistente
        && codigo
    ) {
        alert(
            "Este material ya está agregado al pedido actual."
        );

        return;
    }

    pedido.push(
        {
            codigo_sap:
                codigo,

            descripcion:
                descripcionMaterial.value,

            origen:
                origen.value,

            unidad_medida:
                unidadMedida.value,

            existencia:
                Number(
                    existencia.value
                    || 0
                ),

            consumo_promedio:
                Number(
                    consumoPromedio.value
                    || 0
                ),

            tiempo_entrega:
                Number(
                    tiempoEntrega.value
                    || 0
                ),

            cantidad_pedir:
                Number(
                    cantidadPedir.value
                    || 0
                ),

            valor_unitario:
                Number(
                    valorUnitario.value
                    || 0
                ),

            observaciones:
                observaciones.value.trim(),

            origen_dato:
                "SAP",
        }
    );

    pintarPedido();

    cerrarVentana();
}


function pintarPedido() {
    if (
        !pedido.length
    ) {
        tablaPedido.innerHTML =
            `
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

        generarPedido.disabled =
            true;

        return;
    }

    tablaPedido.innerHTML =
        pedido
            .map(
                (
                    item,
                    indice
                ) => `
                    <tr>

                        <td>
                            ${escaparHtml(
                                item.codigo_sap
                                || "CREAR"
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
                            ${formatearNumero(
                                item.cantidad_pedir
                            )}
                        </td>

                        <td>
                            $${formatearNumero(
                                item.valor_unitario
                            )}
                        </td>

                        <td>

                            <button
                                type="button"
                                class="boton-eliminar"
                                data-indice="${indice}"
                            >
                                ✕
                            </button>

                        </td>

                    </tr>
                `
            )
            .join("");

    document
        .querySelectorAll(
            ".boton-eliminar"
        )
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

    contadorMateriales.textContent =
        `${pedido.length} ${
            pedido.length === 1
            ? "material"
            : "materiales"
        }`;

    generarPedido.disabled =
        false;
}


if (
    buscarMaterial
) {
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


if (
    cerrarModal
) {
    cerrarModal.addEventListener(
        "click",
        cerrarVentana
    );
}


if (
    cerrarFondo
) {
    cerrarFondo.addEventListener(
        "click",
        cerrarVentana
    );
}


if (
    cancelarModal
) {
    cancelarModal.addEventListener(
        "click",
        cerrarVentana
    );
}


if (
    formAgregarMaterial
) {
    formAgregarMaterial.addEventListener(
        "submit",
        agregarMaterialAlPedido
    );
}


if (
    limpiarPedido
) {
    limpiarPedido.addEventListener(
        "click",
        function () {

            if (
                !pedido.length
            ) {
                return;
            }

            const confirmar =
                confirm(
                    "¿Desea eliminar todos los materiales del pedido?"
                );

            if (
                !confirmar
            ) {
                return;
            }

            pedido = [];

            pintarPedido();

        }
    );
}


if (
    generarPedido
) {
    generarPedido.addEventListener(
        "click",
        function () {

            alert(
                "El siguiente paso será guardar la solicitud y generar el formato corporativo."
            );

        }
    );
}