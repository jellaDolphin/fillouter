/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

$(function () {
  $("#dialog-menu").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
  });
  $("#menu-template").button();
  $("#menu-setting").button();
  $("#menu-about").button();

  $("#dialog-template").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
  });

  var dialog_setting;
  dialog_setting = $("#dialog-setting").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    buttons: {
      Connect: function () {},
      Cancel: function () {
        dialog_setting.dialog("close");
      },
    },
  });

  var dialog_save;
  dialog_save = $("#dialog-save").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    buttons: {
      Save: function () {},
      Cancel: function () {
        dialog_save.dialog("close");
      },
    },
  });
  //	$( "#button-save" ).button();
  //	$( "#button-save-cancel" ).button();

  $("#dialog-about").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
  });
  // next add the onclick handler
  $("#menu").click(function () {
    $("#dialog-menu").dialog("open");
    return false;
  });

  $("#menu-template").click(function () {
    $("#dialog-menu").dialog("close");
    $("#dialog-template").dialog("open");
    return false;
  });

  $("#menu-setting").click(function () {
    $("#dialog-menu").dialog("close");
    $("#dialog-setting").dialog("open");
    return false;
  });

  $("#menu-about").click(function () {
    $("#dialog-menu").dialog("close");
    $("#dialog-about").dialog("open");
    return false;
  });

  $("#save-template").click(function () {
    $("#dialog-save").dialog("open");
    return false;
  });

  $(document).on("click", ".delete-section", function () {
    $(this)
      .parents(".ui-state-default")
      .map(function () {
        $(this).remove();
      });
    return false;
  });

  var field_value;
  $(document).on("click", ".insert-field-value", function () {
    field_value = $(this).parents(".tr-field").find(".field-value").val();
    insertFieldValue(field_value);
  });

  $(document).on("click", ".delete-field", function () {
    $(this)
      .parents(".tr-field")
      .map(function () {
        $(this).remove();
      });
    return false;
  });

  $(document).on("click", ".add-field", function () {
    $(this)
      .parents(".table-section")
      .map(function () {
        var $lastRow = $(this).find(".tr-field").last();
        var $newRow = $lastRow.clone();

        $(this).find(".tbody-field").append($newRow);
      });
    return false;
  });

  $("#button-insert-car").click(function () {
    var $li_section = $(".ui-state-default").last().clone();
    $("#sortable").append($li_section);
    currentSectionNo++;
    $li_section.find(".button-no").val("#" + currentSectionNo);

    $li_section.find(".delete-section").click(function () {
      $(this)
        .parents(".ui-state-default")
        .map(function () {
          $(this).remove();
        });
      return false;
    });
    return false;
  });

  $("#sortable").sortable({
    placeholder: "ui-state-highlight",
  });
});

// eslint-disable-next-line no-undef
const gval = require("../../global.json");
var templateOptions = "<option value=''>Select a template</option>";
var genericTextsOptions = "<option value=''>Select text</option>";
var productListOptions = "<option value=''>Select product</option>";
var productTextsOptions = "<option value=''>Select product text</option>";
var i;
var proId = "";
var downfilesURI = [];
var downfilesBuf = [];
var currentSectionNo = 1;
//var months = [{"jan":0}, {"feb":1}, {"mar":2}, {"apr":3}, {"may":4}, {"jun":5}, {"jul":6}, {"aug":7}, {"sep":8}, {"oct":9}, {"nov":10}, {"dec":11}];
var months = { jan: 0, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

//my var
var jsonObject = {};
var currentObject = {};
var g_template_id = 1;
var g_fname = "template1.docx";
var oribuf = {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // eslint-disable-next-line no-undef
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }
    //json, docx download;
    initDownSet();

    for (i in gval.Templates) {
      templateOptions = templateOptions + '<option value="' + i + '">' + gval.Templates[i].Name + "</option>";
      downfilesURI[i] = gval.Templates[i].URI;
    }

    for (i in gval.GenericTexts) {
      genericTextsOptions =
        genericTextsOptions + '<option value="' + i + '">' + gval.GenericTexts[i].Name + "</option>";
    }

    for (i in gval.ProductList) {
      if (i == "0") proId = gval.ProductList[i].Id;
      productListOptions = productListOptions + '<option value="' + i + '">' + gval.ProductList[i].Name + "</option>";
    }

    for (i in gval.ProductTexts) {
      if (proId == gval.ProductTexts[i].Id)
        productTextsOptions =
          productTextsOptions + '<option value="' + i + '">' + gval.ProductTexts[i].Name + "</option>";
    }

    //context
    document.getElementById("select-template").innerHTML = templateOptions;
    // document.getElementById("texts").innerHTML = genericTextsOptions;
    // document.getElementById("products").innerHTML = productListOptions;
    // document.getElementById("product_texts").innerHTML = productTextsOptions;

    // Assign event handlers and other initialization logic.
    document.getElementById("app-body").style.display = "flex";

    //my actions
    document.getElementById("select-template").onchange = insertTemplates;
    // document.getElementById("texts").onchange = insertGenericTexts;
    // document.getElementById("products").onchange = selectProducts;
    // document.getElementById("product_texts").onchange = insertProducts;
    // document.getElementById("id-datepicker-1").onpointerleave = selectProducts;
    document.getElementById("menu").onclick = menuClick;
    document.getElementById("add-template").onclick = addTemplate;
    document.getElementById("save-template").onclick = saveTemplate;
    document.getElementById("button-insert-car").onclick = addCar;
    document.getElementById("button-insert-motorbike").onclick = addBike;

    // let  nodelist = document.querySelectorAll(".insert-field-value");
    // let length = nodelist.length;
    // for (let i = 0 ; i < length ; i++){
    //   nodelist[i].onclick = insertFieldValue;
    // }
    $(document).ready(function () {
      $("#submit").click(function () {
        sendFile();
      });
    });
  }
});

//OK good down file to base64
/*$.post( gval.serverURI, {fpath:"template2.docx"}, function( data ) {
  // updateStatus("RECVDATA initSet success !!!!");
  context.document.body.insertFileFromBase64(data.replace(/^.+,/, ""), Word.InsertLocation.replace);
});*/
function updateStatus(message) {
  var statusInfo = $("#status");
  statusInfo[0].innerHTML += message + "<br/>";
}

function sendFile() {
  Office.context.document.getFileAsync("compressed", { sliceSize: 100000 }, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      // Get the File object from the result.
      var myFile = result.value;
      var state = {
        file: myFile,
        counter: 0,
        sliceCount: myFile.sliceCount,
      };
      getSlice(state);
    } else {
      updateStatus(result.status);
    }
  });
}

function getSlice(state) {
  state.file.getSliceAsync(state.counter, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      sendSlice(result.value, state);
    } else {
      updateStatus(result.status);
    }
  });
}

function sendSlice(slice, state) {
  var data = slice.data;
  if (data) {
    var buf = btoa(data);
    closeFile(state);

    currentObject["buf"] = buf;

    $.post(
      gval.serverURI,
      { jsonData: JSON.stringify(currentObject) },
      function (returnData) {
        if (returnData["data"] == "success") {
          // updateStatus(" ::update SUCCESS!!! " + returnData["buf"][0] + returnData["buf"][1] + returnData["buf"][2]);
          jsonObject["templates"][g_template_id]["buf"] = returnData["buf"];
          oribuf[g_template_id] = data;
        }
      },
      "json"
    );
  }
}

function closeFile(state) {
  state.file.closeAsync(function (result) {
    // eslint-disable-next-line no-empty
    if (result.status == "succeeded") {
    } else {
      // updateStatus("File couldn't be closed.");
    }
  });
}

export async function initDownSet() {
  return Word.run(async (context) => {
    $.post(
      gval.serverURI,
      { initSet: true },
      function (data) {
        jsonObject = {}; jsonObject = data;
        currentObject["template_id"] = g_template_id;
        currentObject["fname"] = g_fname;

        for (const y in jsonObject["template_contents"]) {
          if (jsonObject["template_contents"][y]["template_id"] == g_template_id) {
            var id = jsonObject["template_contents"][y]["car_id"];
            currentObject[id] = {};
            currentObject[id]["id"] = jsonObject["template_contents"][y]["car_id"];
            currentObject[id]["date"] = jsonObject["template_contents"][y]["created_at"];
            currentObject[id]["cars"] = {};
            currentObject[id]["cars"] = jsonObject["cars"][id];

            currentObject[id]["car_values"] = {};
            for (const cv_id in jsonObject["car_values"]) {
              if (jsonObject["car_values"][cv_id]["car_id"] == id) {
                currentObject[id]["car_values"][cv_id] = jsonObject["car_values"][cv_id];
              }
            }

            insertCar(id);
          }
        }
        // updateStatus(JSON.stringify(currentObject));

        // updateStatus(" ::initSet SUCCESS!!! ");
      },
      "json"
    );

    await context.sync();
  });
}

function insertCar(id) {
  $(function () {
    var $li_section =
      "<li class='ui-state-default' id='section-" + currentSectionNo + "'> \
    <div class='div-section'> \
      <table width='100%' class='table-section'> \
      <tbody> \
        <tr class='tr-section-control'> \
        <td> \
          <input name='button' type='button' class='button-no' value='#" + currentSectionNo + "'> \
          <select name='select' class='select'><option value='1'>" + currentObject[id]["cars"]["title"] + "</option> \
          </select> \
          <img src='../../assets/plus.png' alt='' class='add-field' id='add-field-" + id + "'/> \
        </td> \
        <td width='10%' align='right'><img src='../../assets/close.png' alt='' class='delete-section' id='delete-section-" + id + "'/></td> \
        </tr> \
        <tr class='tr-section-fields'> \
        <td colspan='2'><table width='100%' class='table-fields'> \
          <tbody class='tbody-field'>";
    for (const field in currentObject[id]["car_values"]) {
      var field_id = currentObject[id]["car_values"][field]["field_id"];
      var field_title = jsonObject["fields"][field_id]["title"];

      $li_section +=
        "   <tr class='tr-field'> \
              <td><table width='100%'> \
              <tbody> \
                <tr> \
                <td width='6%'><img  src='../../assets/arrow-left.png' alt='' class='insert-field-value'/></td> \
                <td><input name='field-name' type='text' class='field-name' id='field-name-" + id + "-" + field_id + "' value = '" + "[#" + currentSectionNo + "_" + field_title + "]'></td> \
                <td width='6%' align='right'><img src='../../assets/close.png' alt='' class='delete-field'/></td> \
                </tr> \
                <tr> \
                <td colspan='3'> \
                  <input name='textfield2' type='text' class='field-value' id='field-value-" + id + "-" + field_id + "' value='" + currentObject[id]["car_values"][field]["field_value"] + "'> \
                  </td> \
                </tr> \
                </tbody> \
              </table></td> \
            </tr>";
    }
    $li_section +=
      "   </tbody> \
          </table> \
        </td> \
        </tr> \
      </tbody> \
      </table> \
    </div> \
    </li>";

    $("#sortable").append($li_section);
    currentSectionNo++;
  });
}

function insertTemplates() {
  Word.run(function (context) {
    g_template_id = eval(document.getElementById("select-template").value) + 1;
    g_fname = jsonObject["templates"][g_template_id]["fname"];
    currentObject["template_id"] = g_template_id;
    currentObject["fname"] = g_fname;
    for (const y in jsonObject["template_contents"]) {
      if (jsonObject["template_contents"][y]["template_id"] == g_template_id) {
        var id = jsonObject["template_contents"][y]["car_id"];
        currentObject[id] = {};
        currentObject[id]["id"] = jsonObject["template_contents"][y]["car_id"];
        currentObject[id]["date"] = jsonObject["template_contents"][y]["created_at"];
      }
    }
    // document.getElementById("invisible").style.visibility = "visible";

    context.document.body.insertParagraph(
      jsonObject["templates"][g_template_id]["fname"] + jsonObject["templates"][g_template_id]["buf"].length,
      Word.InsertLocation.end
    );
    if (oribuf[g_template_id])
      context.document.body.insertParagraph(oribuf[g_template_id], Word.InsertLocation.replace);
    else
      context.document.body.insertFileFromBase64(
        jsonObject["templates"][g_template_id]["buf"].replace(/^.+,/, "") + "KCG",
        Word.InsertLocation.replace
      );

    context.document.body.insertParagraph(
      jsonObject["templates"][g_template_id]["fname"] + jsonObject["templates"][g_template_id]["buf"].length,
      Word.InsertLocation.end
    );
    return context.sync();
  });
}

function menuClick(e) {
  Word.run(function (context) {
    var currentSelection = context.document.getSelection();

    currentSelection.insertText("menu Clicked ", Word.InsertLocation.end);
    return context.sync();
    // window.alert("menuClicked");
    // console.log("menuClicked");
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function addTemplate(e) {
  Word.run(function (context) {
    console.log("addTemplate");
    var currentSelection = context.document.getSelection();
    currentSelection.insertText("add Template ", Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function saveTemplate(e) {
  Word.run(function (context) {
    console.log("saveTemplate");
    var currentSelection = context.document.getSelection();
    currentSelection.insertText("save Template ", Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function addCar(e) {
  Word.run(function (context) {
    console.log("addCar clicked");
    var currentSelection = context.document.getSelection();
    currentSelection.insertText("add Car ", Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function addBike(e) {
  Word.run(function (context) {
    console.log("addBike clicked");
    var currentSelection = context.document.getSelection();
    currentSelection.insertText("add MotorBike ", Word.InsertLocation.end);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

function insertFieldValue(value) {
  Word.run(function (context) {
    console.log("addField clicked.");
    var currentSelection = context.document.getSelection();
    currentSelection.clear();
    currentSelection.insertText(" " + value, Word.InsertLocation.start);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
