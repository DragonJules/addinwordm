/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

var liste;


Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    document.getElementById("registerclass").onclick = registerclass;

    document.getElementById("registerinput").onclick = registerinput;
  }
});

function registerinput() {
  var input = document.getElementById("listinput").value;

  liste = input;

  var r = document.getElementById("register");
  var e = document.getElementById("error");
  e.style.visibility = "hidden";
  r.style.visibility = "hidden";
  

}

export async function run() {
  return Word.run(async (context) => {
    var e = document.getElementById("error");

    if(!liste) return e.style.visibility = "visible";

    var nouvChain = liste.replace(/-/gi, "\v")

    const paragraph = context.document.body.insertParagraph(nouvChain, Word.InsertLocation.end);

    await context.sync();
  });
}

export async function registerclass() {
  return Word.run(async (context) => {

    var r = document.getElementById("register");

      r.style.visibility = "visible";

    await context.sync();

    var e = document.getElementById("error");
    e.style.visibility = "hidden";
  });
}
