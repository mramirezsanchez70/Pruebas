/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import { CargaXML_PDF } from "./Carga";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {

    // UI
    const sideload = document.getElementById("sideload-msg");
    if (sideload) sideload.style.display = "none";

    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    // 🔥 BOTÓN 1 → Carga XML
    const btnCarga = document.getElementById("btnCarga");
    if (btnCarga) {
      btnCarga.addEventListener("click", CargaXML_PDF);
    }
  }
});