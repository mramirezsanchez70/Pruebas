/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */


import { DatosXML } from "./ExtraerXML";

/**
 * 🔥 FUNCIÓN PRINCIPAL QUE SE LLAMA DESDE taskpane
 */
export function CargaXML_PDF() {

  const inputFile = crearInputArchivo();

  // 🔹 Limpia selección previa
  inputFile.value = "";

  // 🔹 Abre selector de archivos
  inputFile.click();

  // 🔥 IMPORTANTE: evitamos duplicar eventos
  inputFile.onchange = async () => {

    const files = inputFile.files;

    // 🔹 Validación básica
    if (!files || files.length === 0) {
      mostrarModal("No seleccionaste archivos, debes seleccionar un xml con su pdf para su comparativa");
      return;
    }

    const archivos = Array.from(files);

    // 🔹 Contadores
    let contadorXML = 0;
    let contadorPDF = 0;

    let xmlFile: File | null = null;
    let pdfFile: File | null = null;

    // 🔹 Detecta tipo de archivos
    for (const file of archivos) {

      const ext = file.name.split(".").pop()?.toLowerCase();

      if (ext === "xml") {
        contadorXML++;
        xmlFile = file;
      } else if (ext === "pdf") {
        contadorPDF++;
        pdfFile = file;
      }
    }

    // ✅ MÁS DE 2 ARCHIVOS
    if (archivos.length > 2) {
      mostrarModal("Selección incorrecta: para su comparativa, debe seleccionar solo 2 archivo, un xml y su pdf");
      inputFile.value = "";
      return;
    }

    // ✅ SOLO 1 ARCHIVO
    if (archivos.length === 1) {
      mostrarModal("Debe seleccionar un xml y su pdf para su comparativa");
      inputFile.value = "";
      return;
    }

    // ✅ EXACTAMENTE 2 ARCHIVOS
    if (archivos.length === 2) {

      // ❌ Dos XML o dos PDF
      if (contadorXML === 2 || contadorPDF === 2) {
        mostrarModal("Selección incorrecta: para su comparativa, debe seleccionar solo el xml con su pdf");
        inputFile.value = "";
        return;
      }

      // ✅ 1 XML + 1 PDF
      if (contadorXML === 1 && contadorPDF === 1 && xmlFile) {

        // 🔥 ENVÍA EL XML A TU FUNCIÓN
        await DatosXML(xmlFile);
        inputFile.value = "";
        return;

      }
    }
  };
}

/**
 * 🔹 Muestra modal con mensaje
 */
function mostrarModal(mensaje: string) {

  const modal = document.getElementById("modalError") as HTMLElement;
  const texto = document.getElementById("modalMensaje") as HTMLElement;
  const btnCerrar = document.getElementById("btnCerrarModal") as HTMLElement;

  if (texto) texto.textContent = mensaje;
  if (modal) modal.style.display = "block";

  // 🔹 Cerrar modal
  if (btnCerrar) {
    btnCerrar.onclick = () => {
      modal.style.display = "none";
    };
  }
}


/**
 * 🔹 Crea input tipo file oculto
 */
function crearInputArchivo(): HTMLInputElement {

  let input = document.getElementById("fileInput") as HTMLInputElement;

  if (!input) {
    input = document.createElement("input");
    input.type = "file";
    input.accept = ".xml,.pdf";
    input.multiple = true;
    input.style.display = "none";
    input.id = "fileInput";

    document.body.appendChild(input);
  }

  return input;
}