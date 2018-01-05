$(document).ready(function(){
		$('#sigplace').hide();
		$('#sigsection').show();
		$('#sig_2').hide();
		$('#sig_3').hide();
		$('#chkSig').attr('checked', false);
		$('#chkSig').change(function(){
				if ($('#chkSig').prop('checked')) {
					$('#sig_2').show();
					$('#sig_3').show();
					$('#name_first').focus();
				} else {
					$('#sig_2').hide();
					$('#sig_3').hide();
				}

			});
		$('#btnOK').clicked(function(){

			});
	});
