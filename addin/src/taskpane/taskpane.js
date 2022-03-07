/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
var templateOptions = "<option value=''>Select a template</option>";
var carOptions = "<option value=''>Select a car</option>";
var productListOptions = "<option value=''>Select product</option>";
var productTextsOptions = "<option value=''>Select product text</option>";
var i;
var proId = "";
var downfilesBuf = [];
var g_car_id = 1;
var template_status = "";
var g_type_atoi = { car: 1, bike: 2 };
//var months = [{"jan":0}, {"feb":1}, {"mar":2}, {"apr":3}, {"may":4}, {"jun":5}, {"jul":6}, {"aug":7}, {"sep":8}, {"oct":9}, {"nov":10}, {"dec":11}];
var months = { jan: 0, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

//my var
var jsonObject = {};
var currentObject = {};
var g_template_id = 1;
var g_fname = "template1.docx";
var max_field_id = 1;
var max_car_value_id = 1;
var max_car_id = 1;
var g_default_fields = {};
var g_template_names = {};
var g_dir_names = {};
var g_current_dir = "";
var g_current_dir_option = "";
var g_client_token;
var g_server_token ="";
var g_server_token="";g_token = "";
var g_connected = false;
var g_cc_type = 1;
var g_is_new = false;
var g_new_template = 0;

$(function () {
  $("#dialog-menu").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
  });
  $("#menu-template").button();
  $("#menu-setting").button();
  $("#menu-about").button();

  $("#dialog-newfolder").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    maxWidth: 150,
    maxHeight: 300,
    buttons: {
      Save: function () {
        $(this).dialog("close");
        var name = $("#newfolder").val();
        var req = "?newfolder=" + name;
        var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              $("#dialog-setting").dialog("close");
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
              
          }, 2000);

        $.get(
          gval.serverURI + req + g_token,
          {},
          function (data) {
            // updateStatus("newfolder result:" + JSON.stringify(data));
            if (data["status"] == "success") {

              clearTimeout(myVar);
              // if (type == "file")
              var max_id;
              for (var id in g_dir_names) {
                if (max_id < id) max_id = id;
              }
              max_id++;
              g_dir_names[max_id] = name;
              insertFileItem_js("dir", max_id, name, "close");
              insertFileItem_js("dir", max_id, name, "");
              // add option
              // if (g_current_dir_option != "") $("#select-template").append(option);
              // updateStatus("newfolder create:-------" + g_current_dir_option);
              if (g_current_dir_option == "") {
                //$("#select-template option:first").insertAfter(new Option("[" + name + "]"));
                $("#select-template option").eq(1).before($("<option>[" + name + "]</option>"));
              }
            }
            // updateStatus(" delete file g_template_id, id " + g_template_id + " " + id);
          },
          "json"
        );
      },
      Close: function () {
        $(this).dialog("close");
      },
    },
  });
  $("#dialog-template").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    maxWidth: 150,
    maxHeight: 300,
    buttons: [
      {
        id: "btnNewFolder",
        text: "New Folder",
        click: function () {
          $("#dialog-newfolder").dialog("open");
        },
      },
      {
        text: "Close",
        click: function () {
          $(this).dialog("close");
        },
      },
    ],
  });

  $("#dialog-car-name").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    buttons: {
      Save: function () {
        var car_name = $("#input-new-car-name").val();
        // console.log($("#input-new-car-name"));
        $(this).dialog("close");
        // updateStatus("dialog carname:", car_name);

        var t;
        if (g_cc_type == 1) t = "c";
        else t = "b";

        for (var i in currentObject["fields"]) {
          // console.log("car name t, cars_or_bikes:\n", t, currentObject["fields"][i]["cars_or_bikes"]);
          if (currentObject["fields"][i]["cars_or_bikes"] == t) {
            // console.log("car name current:", currentObject["fields"][i]);
            delete currentObject["fields"][i];
          }
        }

        // g_is_new = true;
        insertSection_js(g_cc_type);
        // g_is_new = false;
        // max_car_id++;
        carOptions += '<option value="' + max_car_id + '">' + car_name + "</option>";
        document.getElementById("select-car").innerHTML = carOptions;
        document.getElementById("select-bike").innerHTML = carOptions;
        currentObject["cars"][max_car_id]["title"] = car_name; // get title
        currentObject["cars"][max_car_id]["status"] = "add";
        $(".select").val(max_car_id);

        // console.log("car name currentObject:\n", currentObject);
      },
      Cancel: function () {
        $(this).dialog("close");
      },
    },
  });

  var dialog_overwrite;
  dialog_overwrite = $("#dialog-overwrite").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    maxWidth: 150,
    maxHeight: 300,
    buttons: {
      Yes: function () {
        $(this).dialog("close");
        $("#dialog-save").dialog("close");

        var doc_name = "";
        var full_name = "";
        doc_name = $("#save-doc-name").val();

        if (!doc_name.endsWith(".docx")) {
          doc_name += ".docx";
        }

        if (g_current_dir != "") {
          full_name = g_current_dir + "/" + doc_name;
        } else {
          full_name = doc_name;
        }
        // updateStatus("overwrite---------1------" + full_name);
        currentObject["fname"] = full_name;
        if (currentObject["status"] == "add" || currentObject["status"] == "")
          currentObject["status"] = "replace";

        for (var i in g_template_names) {
          if (full_name == g_template_names[i]) {
            currentObject["id"] = i;
            break;
          }
        }

        sendFile();
        return true;
      },
      No: function () {
        $(this).dialog("close");
        return false;
      },
    },
  });

  var dialog_setting;
  dialog_setting = $("#dialog-setting").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    buttons: [
      {
        id: "btnConnect",
        text: "Connect",
        click: function () {
          
          var myVar;
          myVar = setTimeout(function(){
            
            if(g_server_token == ""){
                $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
                alarm("Can't connect to Server", "");
                $("#div-alarm").css("color", "red");
                dialog_setting.dialog("close");
            }
          }, 2000);

          var disconnect = "";
          if(g_server_token) disconnect = "&disconnect";
          $.get(
            gval.serverURI + "?stoken=" + g_client_token+disconnect,
            {},
            function (returnData) {
              if (returnData["status"] == "success") {
                // updateStatus(" ::token SUCCESS!!! ");
                g_server_token = returnData["plugin_token"];
                // $("#copied_token").val(g_server_token);
                g_connected = true;
                alarm("Connected to API", "");
                $("#div-alarm").css("color", "green");
                getToken();
                // updateStatus("", "");
                initTemplate_js();
                initCars_js();
                initUI();
                clearTimeout(myVar);
                dialog_setting.dialog("close");
              }
              else if(returnData["status"] == "disconnect") {
                //add-action
                alarm("Go Setting & Connect to API", "");
                $("#div-alarm").css("color", "red");
                $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');

                dialog_setting.dialog("close");
                $("#copied-token").val('');
                document.getElementById("select-template").innerHTML = "";
                initUI();
                g_server_token = "";
                clearTimeout(myVar);
              }
              else {
                alarm("Can't connect to Server", "");
                $("#div-alarm").css("color", "red");

                dialog_setting.dialog("close");
              }
            },
            "json"
          );
          // server_token = $("#copied_token").val();
          // server_token = server_token.replace(/ /g, "");
        },
      },
      {
        id: "btnCancel",
        text: "Cancel",
        click: function () {
          dialog_setting.dialog("close");
        },
      },
    ],
  });

  var dialog_save;
  dialog_save = $("#dialog-save").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    maxWidth: 150,
    maxHeight: 300,
    buttons: {
      Save: function () {
        var doc_name = "";
        var full_name = "";
        var is_duplicate = false;
        doc_name = $("#save-doc-name").val();

        if (doc_name == "") return;
        if (!doc_name.endsWith(".docx")) {
          doc_name += ".docx";
        }

        if (g_current_dir != "") {
          full_name = g_current_dir + "/" + doc_name;
        } else {
          full_name = doc_name;
        }

        for (var i in g_template_names) {
          if (full_name == g_template_names[i]) {
            is_duplicate = true;
            break;
          }
        }

        // updateStatus("dialog_save --- fname: " + full_name);
        if (is_duplicate == true) {
          // updateStatus("dialog_save --- is_duplicate: " + is_duplicate);
          $("#dialog-overwrite").dialog("open");
        } else {
          currentObject["fname"] = full_name;
          if (currentObject["status"] == "add") {
            currentObject["id"] = -1;
          }
          $(this).dialog("close");
          sendFile();
        }
      },
      Cancel: function () {
        $(this).dialog("close");
      },
    },
  });

  var dialog_save_confirm;
  dialog_save_confirm = $("#dialog-save-confirm").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
    buttons: {
      Yes: function () {
        $(this).dialog("close");

        if (g_new_template == 1) $("#save-doc-name").val("");
        dialog_save.dialog("open");
      },
      No: function () {
        $(this).dialog("close");
        g_new_template = 1;
        currentObject["status"] = "";
        initTemplate_js();
        // initCars_js();
        // initUI();
      },
      Cancel: function () {
        $(this).dialog("close");
      },
    },
  });

  $("#dialog-about").dialog({
    autoOpen: false,
    modal: true,
    transitionMask: false,
  });

  $("#dialog-alert").dialog({
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
    if (g_new_template == 1) $("#save-doc-name").val("");
    $("#dialog-save").dialog("open");
    return false;
  });

  $("#add-template").click(function () {
    // console.log("add template");
    if (currentObject["status"] == "update") {
      $("#dialog-save-confirm").dialog("open");
      // sendFile();
    } else {
      g_new_template = 1;
      currentObject["status"] = "add";
      initTemplate_js();
      initCars_js();
      initUI();
    }
    return false;
  });

  $(document).on("click", ".file-delete", function () {
    var row = $(this).parent().parent("tr");
    var id = $(this).attr("id");
    // updateStatus("delete file id: " + id);
    var type = id.split("-")[1];
    id = id.split("-")[2];
    var name = "";
    var req = "?";

    if (type == "dir") {
      $(this)
        .parent()
        .parent()
        .children(".dir-name")
        .map(function () {
          name = $(this).html();
          // updateStatus("delete file----------: " + type + " " + id + " " + name);
        });
      req = "?delete_dir=" + name;
    } else {
      name = g_template_names[id];
      req = "?delete_file=" + name;
    }
    // updateStatus("delete file-------name: " + id + " " + name);
    var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);
    $.get(
      gval.serverURI + req + g_token,
      {},
      function (data) {
        if (data["status"] == "success") {
          clearTimeout(myVar);
          // if (type == "file")
          row.remove();
          // updateStatus("file delete--------name: " + type + " " + id + " " + name);
          if (type == "file") {
            for (var i in g_template_names) {
              if (name == g_template_names[i]) {
                delete g_template_names[i];
              }
            }
            $("#select-template option[value=" + id + "]").remove();
          } else {
            
            for (var i in g_dir_names) {
              if (name == g_dir_names[i]) {
                delete g_dir_names[i];
              }
            }

            $("#select-template option").map(function () {
              // updateStatus("delete file----option: " + $(this).val() + $(this).html());
              if (
                $(this)
                  .html()
                  .startsWith("[" + name + "]")
              ) {
                $(this).remove();
              }
            });
          }

          if (type == "file") {
            $("#tbody-saved-files > #tr-file-" + id).remove();
          } else {
            $("#tbody-saved-dirs > #tr-dir-" + id).remove();
          }
          // updateStatus(" delete file g_template_id, id " + g_template_id + " " + id);
          if (g_template_id == id) {
            $("#select-template option").last().prop("selected", true);
            changeTemplate_js();
            // updateStatus(" delete file SUCCESS!!! ");
            // $("#select-template option").val(val).trigger("change");
          }
        }
      },
      "json"
    );
    return true;
  });

  $(document).on("click", ".copy-doc", function () {
    var row = $(this).parent().parent("tr");
    var id = $(this).attr("id");
    // updateStatus("delete file---- id: " + id);
    id = id.split("-")[2];
    var name = "";
    var req = "?";

    // updateStatus("delete file: " + id + " " + name);
    // name = $(this).html();

    $(this)
      .parent()
      .parent()
      .children(".file-name")
      .map(function () {
        name = $(this).html();
        // updateStatus("delete file----------: " + type + " " + id + " " + name);
      });

    // if (type == "dir") req = "?delete_dir=" + name;
    // else
    req = "?copy_file=" + id;
    // updateStatus("copy file----------req: " + req);
    var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);
    $.get(
      gval.serverURI + req + g_token,
      {},
      function (data) {
        //updateStatus("copy file---data: " + JSON.stringify(data));
        if (data["status"] == "success") {
          clearTimeout(myVar);
          var new_id = data["id"];
          var new_name = data["name"];
          var dir_name = "";
          var max_dir_id = 0;
          
          // updateStatus("copy file---1---: " + g_template_names.length);
          // updateStatus("copy file---2---: " + new_id);
          g_template_names[new_id] = new_name;
          // updateStatus("copy file---3---: " + g_template_names.length);
          // updateStatus("copy file---4---: " + g_template_names[new_id]);

          var arr = new_name.split("/");
          if (arr.length > 1) {
            dir_name = arr[0];
            new_name = arr[1];
          }

          // updateStatus("copy doc ---- " + g_current_dir_option + " " + dir_name);
          if (g_current_dir_option == dir_name) {
            $("#select-template").append(new Option(new_name, new_id));
          }

          insertFileItem_js("file", new_id, new_name, "close");
          insertFileItem_js("file", new_id, new_name, "");

          // $("#select-template option[value=" + id + "]").remove();
        }
      },
      "json"
    );
    return true;
  });

  $(document).on("click", ".dir-name", function () {
    // updateStatus("dir-name ---0----");
    $("#btnNewFolder").attr("disabled", true);
    var name = "";
    var fname = "";
    // updateStatus("dir-name ---1----");
    name = $(this).html();
    g_current_dir = name;
    // updateStatus("dir-name : " + name);

    $("#tbody-saved-dirs tr.tr-up").css("display", "block");
    $("#tbody-template-dir tr.tr-up").css("display", "block");
    
    $("#tbody-template-dir tr.tr-file").remove();
    $("#tbody-template-file tr.tr-file").remove();
    $("#tbody-saved-dirs tr.tr-file").remove();
    $("#tbody-saved-files tr.tr-file").remove();
    $("#tbody-saved-files tr.tr-file").remove();
    // updateStatus("dir-name ---2----");
    for (var id in g_template_names) {
      var t_name = g_template_names[id];
      // updateStatus("delete file------t-name: " + t_name);
      if (t_name.startsWith(name + "/")) {
        fname = t_name.split("/")[1];
        insertFileItem_js("file", id, fname, "close");
        insertFileItem_js("file", id, fname, "");
      }
    }
    return true;
  });

  $(document).on("click", ".dir-up", function () {
    // updateStatus("dir up: ");
    $("#btnNewFolder").attr("disabled", false);
    $("#tbody-saved-dirs tr.tr-up").css("display", "none");
    $("#tbody-template-dir tr.tr-up").css("display", "none");
    g_current_dir = "";
    initFileList_js();
    return true;
  });

  $(document).on("click", ".delete-section", function () {
    $(this)
      .parents(".div-section")
      .map(function () {
        $(this).css("display", "none");
        $(this).find(".field-value").val("");
        $(this).find(".field-value").remove();
        // $(this).remove();
      });

    var id = $(this).attr("id");
    var car_type = id.split("-")[1];
    var type = g_type_atoi[car_type];
    // console.log("delete-section g_type_atoi:", g_type_atoi);
    // // console.log("delete-section car_type:", car_type);
    // console.log("delete-section type:", type);

    car_type = car_type[0];
    // updateStatus("delete-section car_type:" + car_type);

    for (field_id in currentObject["fields"]) {
      if (currentObject["fields"][field_id]["cars_or_bikes"] == car_type) {
        if (currentObject["fields"][field_id]["status"] == "add") {
          currentObject["fields"][field_id]["status"] = "ignore";
        } else if (currentObject["fields"][field_id]["status"] != "ignore") {
          currentObject["fields"][field_id]["status"] = "delete";
        }
      }
    }
    // if (currentObject["cars"][type]["status"] == "add") {
    //     currentObject["cars"][type]["status"] = "ignore";
    //   } else {
    //     currentObject["cars"][type]["status"] = "delete";
    //   }
    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }
    // g_is_new = false;
    template_status = "update";
    // updateStatus("delete-section currentObject :\n" + JSON.stringify(currentObject));
    // updateStatus("\n\n");
    return false;
  });

  var field_value;
  $(document).on("click", ".insert-field-value", function () {
    field_value = $(this).parents(".tr-field").find(".field-name").val();
    insertFieldValue(field_value);
    // console.log("insert field value currentObject:\n", currentObject);
  });

  $(document).on("click", ".delete-field", function () {
    var id = $(this).attr("id");
    var ar = id.split("-");

    if (ar.length < 2) {
      $("#dialog-alert").dialog("open");
      return false;
    }

    var field_id = ar[2];

    $(this)
      .parents(".tr-field")
      .map(function () {
        $(this).remove();
      });

    if (currentObject["fields"][field_id]["status"] == "add") {
      currentObject["fields"][field_id]["status"] = "ignore";
    } else {
      currentObject["fields"][field_id]["status"] = "delete";
    }

    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }
    template_status = "update";
    // console.log("deleted-field  currentObject:\n ", currentObject);
    return false;
  });

  $(document).on("click", ".add-field", function () {
    max_car_value_id++;
    max_field_id++;
    var id = $(this).attr("id");
    var car_type = id.split("-")[2];
    var field_title = "new_field";
    var type = g_type_atoi[car_type];

    insertField_js(type, max_field_id, field_title);
    var field_id = max_field_id;

    currentObject["fields"][field_id] = {};
    currentObject["fields"][field_id]["id"] = -1;
    currentObject["fields"][field_id]["fields"] = field_title;
    currentObject["fields"][field_id]["cars_or_bikes"] = car_type[0];
    currentObject["fields"][field_id]["is_default"] = "0";
    currentObject["fields"][field_id]["status"] = "add";

    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }

    // console.log("add-field  currentObject:\n ", currentObject);
    return true;
  });

  $(document).on("change", ".field-value", function () {
    var id = $(this).attr("id");
    var changed_value = $(this).val();
    var ar = id.split("-");

    if (ar.length < 4) {
      $("#dialog-alert").dialog("open");
      return false;
    }
    var value_id = ar[2];
    var default_name = "";
    $(this)
      .parents("tbody")
      .map(function () {
        $(this)
          .find(".field-name")
          .map(function () {
            var name_id = $(this).attr("id");
            var name_ar = name_id.split("-");

            if (name_ar.length < 4) {
              $("#dialog-alert").dialog("open");
              return false;
            }
            name_id = name_ar[2];
            // updateStatus("chang fieldvalue---ids :" + value_id + " " + name_id);
            if (value_id != name_id) return;
            var origin = $(this).val();
            var pattern = /[a-zA-Z]+[a-zA-Z_]+/g;
            var field_name = origin.match(pattern);
            // updateStatus("chang fieldvalue---field_name:" + field_name);

            for (var id in g_default_fields) {
              if (field_name == g_default_fields[id]["fields"]) {
                default_name = field_name;
              }
            }

            if (default_name == "") {
              return false;
            }
          });
      });
    // updateStatus("change fieldvalue---default_name:" + default_name);

    $(this)
      .parents(".table-section")
      .map(function () {
        var car_id = $(this).find(".select").val();
        // updateStatus("change fieldvalue---car_id:\n    " + car_id + default_name + changed_value);
        // console.log("------- car_id:", car_id);
        currentObject["cars"][car_id][default_name] = changed_value;
        if (currentObject["cars"][car_id]["status"] != "add") {
          currentObject["cars"][car_id]["status"] = "update";
        }
      });

    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }
    // updateStatus("change fieldvalue---curObj(cars):" + JSON.stringify(currentObject["cars"]));
    // console.log("change field value : ", currentObject);
    return true;
  });

  $(document).on("focusout", ".field-name", function () {
    var id = $(this).attr("id");
    var orign = $(this).val();
    var ar = id.split("-");

    if (ar.length < 4) {
      $("#dialog-alert").dialog("open");
      return false;
    }

    var field_view_name;
    var field_id = ar[2];
    var pattern = /[a-zA-Z]+[a-zA-Z_]+/g;
    var field_name = orign.match(pattern);
    // updateStatus("focusout.field-name : " + field_name);
    for (var i in g_default_fields) {
      if (field_name == g_default_fields[i]["fields"]) {
        return false;
      }
    }

    // if (field_id in currentObject["fields"] && !(field_id in g_default_fields)) {
    // } else {
    //   return false;
    // }
    field_name = field_name[0];
    if (ar[3] == "car") field_view_name = "[#1_" + field_name + "]";
    else field_view_name = "[#2_" + field_name + "]";

    $(this).val(field_view_name);
    if (currentObject["fields"][field_id]["fields"] == field_name) return false;

    currentObject["fields"][field_id]["fields"] = field_name;

    if (currentObject["fields"][field_id]["id"] != -1) {
      currentObject["fields"][field_id]["status"] = "update";
    }

    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }
    // console.log("focusout field-name :\n ", currentObject);
    // console.log("focusout field_name currentObject:\n", currentObject);
    return true;
  });

  $(document).on("focus", ".field-name", function () {
    var id = $(this).attr("id");
    var origin = $(this).val();
    var ar = id.split("-");
    // var field_id = ar[2];
    // updateStatus("focus.field-name ar: " + ar);
    if (ar.length < 4) {
      $("#dialog-alert").dialog("open");
      return false;
    }

    var pattern = /[a-zA-Z]+[a-zA-Z_]+/g;
    var field_name = origin.match(pattern);
    // updateStatus("focus.field-name : " + field_name);
    for (var i in g_default_fields) {
      if (field_name == g_default_fields[i]["fields"]) {
        return false;
      }
    }
    $(this).val(field_name);

    // if (field_id in currentObject["fields"] && !(field_id in g_default_fields)) {
    //   field_name = currentObject["fields"][field_id]["fields"];
    // } else {
    //   return false;
    // }
    return true;
  });

  $(document).on("change", ".select", function () {
    var id = $(this).attr("id");
    // console.log("id:", id);
    g_car_id = $(this).val();
    // console.log("g_car_id", g_car_id);
    var ar = id.split("-");
    // // console.log("change select id: ", ar);
    var select_dom = $(this);
    var default_name = "";
    // updateStatus("change.select g_default_fields: " + g_default_fields);
    // updateStatus("change.select g_car_id: " + g_car_id);
    var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);
    $.get(
      gval.serverURI + "?car=" + g_car_id + g_token,
      {},
      function (data) {
        clearTimeout(myVar);
        jsonObject = {};
        jsonObject = data;

        if (ar.length < 2) {
          $("#dialog-alert").dialog("open");
          return false;
        }

        currentObject["cars"][g_car_id] = {};
        currentObject["cars"][g_car_id] = jsonObject;
        // console.log("change select jsonObjec:\n", jsonObject);
        // console.log("change select g_default_fields:\n", g_default_fields);
        // updateStatus("change.select ----jsonObject : " + JSON.stringify(jsonObject));
        // get field name;
        var father;
        var sel_id;
        $(select_dom)
          .parents(".div-section")
          .map(function () {
            var father = $(this);
            $(this)
              .find(".field-name")
              .map(function () {
                // father = $(this).parent()[0];
                // console.log("\n change select id:\n", $(this).attr("id"));
                ar = $(this).attr("id").split("-");
                // updateStatus("change.select ----name_id : " + $(this).attr("id"));
                sel_id = ar[2];

                var origin = $(this).val();
                var pattern = /[a-zA-Z]+[a-zA-Z_]+/g;
                var field_name = origin.match(pattern);
                // updateStatus("change.select ----field_name : " + field_name);
                default_name = "";
                for (var i in g_default_fields) {
                  if (field_name == g_default_fields[i]["fields"]) {
                    default_name = field_name;
                    // updateStatus("change.select ----default_name : " + g_default_fields[i]["fields"]);
                    break;
                  }
                }
                if (default_name == "") {
                  return false;
                }
                // updateStatus("change.select ------field_name--- : " + field_name);
                // updateStatus("change.select ------father--- : " + father);

                father.find(".field-value").map(function () {
                  var val_id = $(this).attr("id");
                  // updateStatus("change.select ------val_id : " + val_id);
                  val_id = $(this).attr("id").split("-")[2];
                  if (sel_id == val_id) {
                    $(this).val(jsonObject[default_name]);
                    // updateStatus("change.select ----setvalue : " + $(this).val());
                  }
                  // updateStatus("change.select ------father--- : " + father.attr("id"));
                  // if ($(this).parent() == father)
                });
              });
          });
        // updateStatus("\n\n\n");
        // console.log("change select currentObject:\n", currentObject);
      },
      "json"
    );

    // updateStatus(" init car detail SUCCESS!!! ");
  });

  $("#button-insert-motorbike").click(function () {
    if ($("#div-section-bike").css("display") == "none") {
      $("#div-section-bike").css("display", "block");
      if ($("#div-section-car").css("display") != "none") $("#separator").css("display", "block");
      g_is_new = true;
      insertSection_js(2);
      g_is_new = false;
    } else {
      // g_cc_type = 2;
      // $("#dialog-car-name").dialog("open");
    }
    // console.log("button-insert-motorbike currentObject:\n", currentObject);
  });

  $("#button-insert-car").click(function () {
    if ($("#div-section-car").css("display") == "none") {
      $("#div-section-car").css("display", "block");
      g_is_new = true;
      insertSection_js(1);
      g_is_new = false;
    } else {
      // g_cc_type = 1;
      // $("#dialog-car-name").dialog("open");
    }
    // console.log("button-insert-car currentObject:\n", currentObject);
  });

  $("#button-fillout").click(function () {
    $(".tr-field").map(function () {
      var key = $(this).find(".field-name").val();
      var value = $(this).find(".field-value").val();
      replaceWordText(key, value);
    });
    // $(function () {
    // });
  });

  $("#button-reset").click(function () {
    // $(function () {
    //   $(".field-value").map(function () {
    //     $(this).val("");
    //   });
    // });
    insertTemplates();
  });

  $(document).on("click", "#select-template", function (event) {
    if ($("#select-template option").length < 2 && $(this).val().startsWith("..")) {
      g_current_dir_option = "";
      initTemplate_js();
    } else if (currentObject["status"] == "update") {
      event.stopPropagation();
      $("#dialog-save-confirm").dialog("open");
      // sendFile();
    }
    return true;
  });

  $(document).on("input", "#save-doc-name", function (event) {
    var value = $(this).val();
    for (var i in g_template_names) {
      if (value == g_template_names[i]) {
        $(this).css("border-color", "red");
        return true;
        //$("#name-alarm").innerHTML("Not available");
      } else {
        //$("#name-alarm").innerHTML("");
      }
    }
    $(this).css("border-color", "black");
    return true;
  });
});

function insertField_js(type, id, fd_name, readonly = 0) {
  var str_type = { 1: "car", 2: "bike" };
  var id_suffix = id + "-" + str_type[type];
  var fd_name_value = "[#" + type + "_" + fd_name + "]";
  var closeButton = "";
  if (readonly == 1) {
    readonly = "readonly";
  } else {
    readonly = "";
    closeButton = "<img src='../../assets/close.png' alt='' class='delete-field' id='delete-field-" + id + "'/>";
  }

  var tr_field =
    "   <tr class='tr-field'> \
      <td><table width='100%'> \
      <tbody> \
        <tr> \
        <td width='6%'><img  src='../../assets/arrow-left.png' alt='' class='insert-field-value'/></td> \
        <td><input name='field-name' type='text' class='field-name' id = 'field-name-" +
    id_suffix +
    "' value='" +
    fd_name_value +
    "' " +
    readonly +
    "></td> \
        <td width='6%' align='right'>" +
    closeButton +
    "</td> \
        </tr> \
        <tr> \
        <td colspan='3'> \
          <input name='textfield2' type='text' class='field-value' id='field-value-" +
    id_suffix +
    "'> \
          </td> \
        </tr> \
        </tbody> \
      </table></td> \
    </tr>";

  if (type == 1) {
    $("#tbody-field-car").append(tr_field);
  } else {
    $("#tbody-field-bike").append(tr_field);
  }
}

function insertSection_js(type) {
  var today = new Date();
  var date = today.getFullYear() + "-" + (today.getMonth() + 1) + "-" + today.getDate();
  var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
  var dateTime = date + " " + time;
  max_car_id++;
  // console.log("insertSection_js -------- g_is_new:\n", g_is_new);
  if (!g_is_new) {
    currentObject["cars"][max_car_id] = {};
    currentObject["cars"][max_car_id]["id"] = -1;
    currentObject["cars"][max_car_id]["title"] = ""; // get title
    currentObject["cars"][max_car_id]["brand"] = ""; // get title
    currentObject["cars"][max_car_id]["year"] = ""; // get title
  }

  if (type == 1) {
    $("#tbody-field-car .tr-field").remove();
  } else {
    $("#tbody-field-bike .tr-field").remove();
  }

  for (var id in g_default_fields) {
    if (g_default_fields[id]["cars_or_bikes"] == "c" && type != 1) continue;
    else if (g_default_fields[id]["cars_or_bikes"] == "b" && type != 2) continue;
    // console.log("g_default_fields's type:", g_default_fields[id]["cars_or_bikes"]);

    max_field_id++;
    var fd_name = g_default_fields[id]["fields"];
    // updateStatus("insertSection----type, id, fd_name: " + type + " " + id + " " + fd_name);
    insertField_js(type, id, fd_name, 1);
    currentObject["fields"][max_field_id] = {};
    currentObject["fields"][max_field_id]["id"] = -1;
    currentObject["fields"][max_field_id]["fields"] = g_default_fields[id]["fields"];
    currentObject["fields"][max_field_id]["cars_or_bikes"] = g_default_fields[id]["cars_or_bikes"];
    currentObject["fields"][max_field_id]["is_default"] = "1";
    currentObject["fields"][max_field_id]["status"] = "add";
    if (currentObject["status"] != "add") {
      currentObject["status"] = "update";
    }
  }
}

function insertFileItem_js(type, id, name, readonly) {
  // type : file or dir
  // readonly : close or empty
  var closeButton = "";
  var copyButton = "";
  var icon = "";
  var td_name = "";
  if (readonly == "close") {
    closeButton =
      "<td width='10%'><img src='../../assets/close.png' alt='' class='file-delete' id='delete-" +
      type +
      "-" +
      id +
      "'/></td>";
    copyButton =
      "<td width='10%'><img src='../../assets/copy.png' alt='' class='copy-doc' id='copy-file-" + id + "'/></td>";
  }

  var tr_field;
  if (type == "dir") {
    icon = '<td width="10%"><img src="../../assets/folder.png" alt="" class="icon-folder"/></td>';
    copyButton = "<td width='10%'>&nbsp;</td>";
    td_name = '<td class="dir-name">' + name + "</td>";
  } else {
    icon = '<td width="10%">&nbsp;</td>';
    td_name = '<td class="file-name">' + name + "</td>";
  }
  
  if (readonly == "close") {
    tr_field = '<tr class="tr-file">' + icon + td_name + copyButton + closeButton + "</tr>";
  } else {
    if (type == "file") {
      tr_field = '<tr class="tr-file" id="tr-file-' + id + '">' + icon + td_name + copyButton + closeButton + "</tr>";
    } else {
      tr_field = '<tr class="tr-file" id="tr-dir-' + id + '">' + icon + td_name + copyButton + closeButton + "</tr>";
    }
  }


  if (type == "file") {
    if (readonly == "close") $("#tbody-template-file").append(tr_field);
    else $("#tbody-saved-files").append(tr_field);
  } else {
    if (readonly == "close") $("#tbody-template-dir").append(tr_field);
    else $("#tbody-saved-dirs").append(tr_field);
  }
}

function initFileList_js() {
  // console.log("insertSection_js -------- g_is_new:\n", g_is_new);
  $("#tbody-template-dir tr.tr-file").remove();
  $("#tbody-template-file tr.tr-file").remove();
  $("#tbody-saved-dirs tr.tr-file").remove();
  $("#tbody-saved-files tr.tr-file").remove();

  for (var id in g_dir_names) {
    insertFileItem_js("dir", id, g_dir_names[id], "close");
    insertFileItem_js("dir", id, g_dir_names[id], "");
  }

  for (var id in g_template_names) {
    if (g_template_names[id].split("/").length == 1) {
      insertFileItem_js("file", id, g_template_names[id], "close");
      insertFileItem_js("file", id, g_template_names[id], "");
    }
  }

  // jsonObject["dirs"] = {1:"aaa",2:"bbb",3:"ccc"};
}

function updateStatus(message, act = "add") {
  var statusInfo = $("#status");
  if (act != "add") statusInfo[0].innerHTML = message + "<br/>";
  else statusInfo[0].innerHTML += message + "<br/>";
}

function alarm(message, act = "add") {
  var statusInfo = $("#div-alarm");
  if (act != "add") statusInfo[0].innerHTML = message + "<br/>";
  else statusInfo[0].innerHTML += message + "<br/>";
}

function initTemplate_js() {
  
  var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);

  $.get(
    gval.serverURI + "?templates" + g_token,
    {},
    function (data) {

      clearTimeout(myVar);
      if(data["status"] != "success")
        return;
      jsonObject = {};
      jsonObject = data;
      // console.log("initTemplate jsonObject:\n", jsonObject);
      g_default_fields = jsonObject["fields"];
      // updateStatus("--------g_default_fields:" + JSON.stringify(g_default_fields));

      templateOptions = "<option value=''>Select a template</option>";
      // updateStatus( JSON.stringify(jsonObject["templates"]) );
      g_dir_names = jsonObject["dirs"];

      for (var id in g_dir_names) {
        templateOptions += "<option>[" + g_dir_names[id] + "]</option>";
      }
      // updateStatus("initTemplate------\n" + JSON.stringify(jsonObject));
      var i;
      for (var ii in jsonObject["templates"]) {
        i = jsonObject["templates"][ii]["id"];
        g_template_names[i] = jsonObject["templates"][ii]["file_name"];

        if (g_current_dir_option != "") {
        }

        if (g_template_names[i].search("/") < 0) {
          templateOptions =
            templateOptions + '<option value="' + i + '">' + jsonObject["templates"][ii]["file_name"] + "</option>";
        }
      }
      // updateStatus("initTemplate_js array: " + g_template_names);

      g_template_id = 0;
      document.getElementById("select-template").innerHTML = templateOptions;
      $("#table-setting-success").css("display","block");
      $("#table-setting").css("display","none");

      $("#btnConnect").html("disConnect");
      $("#btnConnect").css('width', '120px');
      $("#btnConnect").css('height', '40px');

      currentObject = {};
      currentObject["buf"] = "";
      currentObject["id"] = -1;
      currentObject["fname"] = "";
      currentObject["date"] = "";
      currentObject["buf"] = "";
      currentObject["cars"] = {};
      currentObject["fields"] = {};

      initFileList_js();
      // console.log("initTemplate g_default_fields:\n", g_default_fields);
      // updateStatus(" init template SUCCESS!!! ");
    },
    "json"
  );
}


function initCars_js() {

  var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);

  $.get(
    gval.serverURI + "?cars" + g_token,
    {},
    function (data) {

      clearTimeout(myVar);
      // var data = '{"1":"Peters car","2":"Vals car","3":"Johns car"}';
      jsonObject = {};
      jsonObject = data; // JSON.parse(data);
      // // console.log("cars:", jsonObject);

      carOptions = "<option value=''>Select a car</option>";
      for (var i in jsonObject) {
        // console.log("car id:", i);
        carOptions = carOptions + '<option value="' + i + '">' + jsonObject[i] + "</option>";
        if (max_car_id < eval(i)) max_car_id = eval(i);
        // console.log("max_car_id:\n", max_car_id);
      }

      g_template_id = 0;
      document.getElementById("select-car").innerHTML = carOptions;
      document.getElementById("select-bike").innerHTML = carOptions;

      // updateStatus(" init car SUCCESS!!! ");
    },
    "json"
  );
}

function initUI() {
  $(".tr-field").map(function () {
    $(this).remove();
  });
  $("#div-section-car").css("display", "none");
  $("#div-section-bike").css("display", "none");
  $("#separator").css("display", "none");
}

$(document).ready(function () {
  // initTemplate_js();
  // initCars_js();
  // initUI();
  g_client_token = generate_token(32);
  $("#new-token").val(g_client_token);

  document.getElementById("app-body").style.display = "flex";
  // document.getElementById("select-template").onchange = insertTemplates;
  document.getElementById("select-template").onchange = changeTemplate_js;
});

// create fields
function changeTemplate_js(event) {
  $("#button-fillout").removeAttr("disabled");
  $("#button-reset").removeAttr("disabled");
  var e = document.getElementById("select-template");
  var pattern = /[a-zA-Z_\s]+/g;
  var origin = e.value;
  var dir_name = origin.match(pattern);
  // updateStatus("changeTemplate--------1111111111----" + dir_name);

  if (origin.endsWith(".docx") < 0) g_current_dir_option = dir_name;
  else g_current_dir_option = "";
  // updateStatus("------\n" + JSON.stringify(g_template_names));
  if (e.value.startsWith("[") == true) {
    templateOptions = "<option>../[" + dir_name + "]</option>";
    for (var i in g_template_names) {
      // updateStatus("changeTemplate-----222222222-------" + g_template_names[i] + " " + g_template_names[i].startsWith(dir_name + "/"));
      if (g_template_names[i].startsWith(dir_name + "/") && g_template_names[i].split("/")[1]) {
        templateOptions += '<option value="' + i + '">' + g_template_names[i].split("/")[1] + "</option>";
      }
    }
    event.stopPropagation();
    document.getElementById("select-template").innerHTML = templateOptions;
  } else if (e.value.startsWith("..") == true) {
    // updateStatus("changeTemplate-----3333333333-------");
    g_current_dir_option = "";
    event.stopPropagation();
    // updateStatus("changeTemplate------------");
    initTemplate_js();
  }

  g_template_id = eval(e.value);
  g_fname = e.options[e.selectedIndex].text;
  $("#save-doc-name").val(g_fname);
  g_new_template = 0;

  // console.log("g_fname:", g_fname);

  currentObject["id"] = g_template_id;
  currentObject["fname"] = g_fname;
  currentObject["cars"] = {};
  // updateStatus("changeTemplate_js --- 1");
  var myVar;
          myVar = setTimeout(function(){
            g_server_token="";g_token = "";
            $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);
  $.get(
    gval.serverURI + "?template=" + g_template_id + g_token,
    {},
    function (data) {
      clearTimeout(myVar);
      // console.log("fields:", data);
      jsonObject = {};
      jsonObject = data; // JSON.parse(data);
      currentObject["fields"] = jsonObject["fields"];
      currentObject["buf"] = jsonObject["buf"];
      currentObject["status"] = "";
      // updateStatus("changeTemplate_js recevied: " + JSON.stringify(currentObject["fields"]));
      // updateStatus("changeTemplate_js recevied: " + currentObject["buf"].length);

      insertTemplates();

      $(".tr-field").map(function () {
        $(this).remove();
      });

      max_field_id = 1;
      for (var id in jsonObject["fields"]) {
        var type = jsonObject["fields"][id]["cars_or_bikes"];
        var field_name = jsonObject["fields"][id]["fields"];
        var is_default = jsonObject["fields"][id]["is_default"];
        // // console.log("field's type:", type);
        if (id > max_field_id) max_field_id = id;
        if (type == "c") {
          type = 1;
          $("#div-section-car").css("display", "block");
        } else {
          type = 2;
          $("#div-section-bike").css("display", "block");
          if ($("#div-section-car").css("display") != "none") $("#separator").css("display", "block");
        }

        var readOnly = 0;
        for (var d_id in g_default_fields) {
          if (field_name == g_default_fields[d_id]["fields"]) {
            readOnly = 1;
          }
        }
        // if (id in g_default_fields) readOnly = 1;

        // updateStatus("changeTemplate-------field_name : " + field_name);
        // for (var default_id in g_default_fields) {
        //   // updateStatus("changeTemplate-------default id : " + default_id);
        //   // updateStatus("changeTemplate-------name : " + g_default_fields[default_id]["fields"]);
        //   // if (field_name == g_default_fields[default_id]["fields"]) id = default_id;
        // }

        insertField_js(type, id, field_name, is_default);
      }
      // console.log("changeTemplate_js currentObject:\n", currentObject);
    },
    "json"
  );
}


function generate_token(length) {
  //edit the token allowed characters
  var a = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890".split("");
  var b = [];
  for (var i = 0; i < length; i++) {
    var j = (Math.random() * (a.length - 1)).toFixed(0);
    b[i] = a[j];
  }
  return b.join("");
}

function getToken() {
  if (g_connected) {
    g_token = "&stoken=" + g_client_token + "&ptoken=" + g_server_token;
  }
  return g_token;
}

function replaceWordText(key, value) {
  Word.run(function (context) {
    // Queue a command to search the document and ignore punctuation.
    var searchResults = context.document.body.search(key, { ignorePunct: true });
    context.load(searchResults);

    // updateStatus(key + value);
    return context.sync().then(function () {
      for (var i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].insertText(value, Word.InsertLocation.replace);
      }

      return context.sync();
    });
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // eslint-disable-next-line no-undef
      // console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }
    Office.addin.setStartupBehavior(Office.StartupBehavior.load);
    //context
    // document.getElementById("select-template").innerHTML = templateOptions;

    // Assign event handlers and other initialization logic.

    //my actions
    // jquery function uesed
    // $(document).ready(function () {});
  }
});

//OK good down file to base64
// $.post( gval.serverURI, {fpath:"template2.docx"}, function( data ) {
//   // updateStatus("RECVDATA initSet success !!!!");
//   context.document.body.insertFileFromBase64(data.replace(/^.+,/, ""), Word.InsertLocation.replace);
// });

function sendFile() {
  // updateStatus("send File called");
  // // console.log("sendFile currentObject:\n", currentObject);
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
      // updateStatus(result.status);
    }
  });
}

function getSlice(state) {
  state.file.getSliceAsync(state.counter, function (result) {
    if (result.status == Office.AsyncResultStatus.Succeeded) {
      sendSlice(result.value, state);
    } else {
      // updateStatus(result.status);
    }
  });
}

function sendSlice(slice, state) {
  var data = slice.data;
  if (data) {
    var buf = btoa(data);
    closeFile(state);

    var today = new Date();
    var date = today.getFullYear() + "-" + (today.getMonth() + 1) + "-" + today.getDate();
    var time = today.getHours() + ":" + today.getMinutes() + ":" + today.getSeconds();
    var dateTime = date + " " + time;

    currentObject["date"] = dateTime;
    currentObject["buf"] = "";
    // updateStatus("\n----Sending file-->");
    // updateStatus("send:");
    //updateStatus(JSON.stringify(currentObject));

    var myVar;
          myVar = setTimeout(function(){
              g_server_token="";g_token = "";
              $("#table-setting-success").css("display","none");
                $("#table-setting").css("display","block");

                $("#btnConnect").html("Connect");
                $("#btnConnect").css('width', '120px');
                $("#btnConnect").css('height', '40px');
              alarm("Can't connect to Server", "");
              $("#div-alarm").css("color", "red");
          }, 2000);

    currentObject["buf"] = buf;
    $.post(
      gval.serverURI,
      { template_save: JSON.stringify(currentObject), stoken: g_client_token, ptoken: g_server_token },
      function (returnData) {
        clearTimeout(myVar);
        //updateStatus(JSON.stringify(currentObject));
        if (returnData["status"] == "success") {
          currentObject["status"] = "";
          initTemplate_js();
          initCars_js();
          initUI();
        } else {
          // updateStatus(" ::error : " + returnData["buf"]);
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

function insertTemplates() {
  Word.run(function (context) {
    // changeTemplate_js();
    // document.getElementById("invisible").style.visibility = "visible";

    context.document.body.insertFileFromBase64(currentObject["buf"], Word.InsertLocation.replace);
    return context.sync();
  });
}

function insertFieldValue(value) {
  Word.run(function (context) {
    // console.log("addField clicked.");
    var currentSelection = context.document.getSelection();
    currentSelection.clear();
    currentSelection.insertText(" " + value, Word.InsertLocation.start);

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    // console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      // console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
