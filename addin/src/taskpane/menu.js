$(function() {
  $("#dialog-menu").dialog({
    autoOpen : false, modal : true, show : "blind", hide : "blind"
  });
	$( "#menu-template" ).button(); 
	$( "#menu-setting" ).button(); 
	$( "#menu-about" ).button(); 

  $("#dialog-template").dialog({
    autoOpen : false, modal : true, show : "blind", hide : "blind"
  });

  var dialog_setting;
  dialog_setting = $("#dialog-setting").dialog({
    autoOpen : false, modal : true, show : "blind", hide : "blind",
	  buttons: {
		Connect: function(){},
        Cancel: function() {
          dialog_setting.dialog("close");
        }
      }
  });

  var dialog_save;
  dialog_save = $("#dialog-save").dialog({
    autoOpen : false, modal : true, show : "blind", hide : "blind",
	  buttons: {
		Save: function(){},
        Cancel: function() {
          dialog_save.dialog("close");
        }
      }
  });
//	$( "#button-save" ).button(); 
//	$( "#button-save-cancel" ).button(); 

  $("#dialog-about").dialog({
    autoOpen : false, modal : true, show : "blind", hide : "blind"
  });
  // next add the onclick handler
  $("#menu").click(function() {
    $("#dialog-menu").dialog("open");
    return false;
  });

  $("#menu-template").click(function() {
    $("#dialog-menu").dialog("close");
    $("#dialog-template").dialog("open");
    return false;
  });

  $("#menu-setting").click(function() {
    $("#dialog-menu").dialog("close");
    $("#dialog-setting").dialog("open");
    return false;
  });

  $("#menu-about").click(function() {
    $("#dialog-menu").dialog("close");
    $("#dialog-about").dialog("open");
    return false;
  });

  $("#save-template").click(function() {
    $("#dialog-save").dialog("open");
    return false;
  });

	$(".delete-section").click(function(){
		$(this).parents(".ui-state-default").map(function(){
			$(this).remove();
		});
		return false;
	});

  // $(".insert-field-value").click(insertFieldValue);

	$(".delete-field").click(function(){
		$(this).parents(".tr-field").map(function(){
			$(this).remove();
		});
		return false;
	});

	$(".add-field").click(function(){
		$(this).parents(".table-section").map(function(){

			var $lastRow = $(this).find(".tr-field").last();
			var $newRow = $lastRow.clone();
			
			$(this).find(".tbody-field").append($newRow);

      // $(this).find(".insert-field-value").click(insertFieldValue);

			$(this).find(".delete-field").click(function(){
				$(this).parents(".tr-field").map(function(){
					$(this).remove();
				});
				return false;
			});

		});
		return false;
	});

	$("#button-insert-car").click(function(){
		var $li_section = $(".ui-state-default").last().clone();
		$("#sortable").append($li_section);

		$li_section.find(".delete-section").click(function(){
			$(this).parents(".ui-state-default").map(function(){
				$(this).remove();
			});
			return false;
		});

    // $div_section.find(".insert-field-value").click(insertFieldValue);

		$li_section.find(".delete-field").click(function(){
			$(this).parents(".tr-field").map(function(){
				$(this).remove();
			});
			return false;
		});

		$li_section.find(".add-field").click(function(){
			$(this).parents(".table-section").map(function(){

				var $lastRow = $(this).find(".tr-field").last();
				var $newRow = $lastRow.clone();

				$(this).find(".tbody-field").append($newRow);

        // $(this).find(".insert-field-value").click(insertFieldValue);

				$(this).find(".delete-field").click(function(){
					$(this).parents(".tr-field").map(function(){
						$(this).remove();
					});
					return false;
				});
			});
			return false;
		});
		return false;
	});
	
	$("#sortable").sortable({
      placeholder: "ui-state-highlight"
    });
});