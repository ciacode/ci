<script src="<?php echo base_url()?>js/jquery-1.8.3.js" type="text/javascript"></script>
<script>
function getfile(){
			$.post("<?php echo site_url()?>/d_excel/download_excel",function(data){
				window.location.href=data;
			});
}
</script>

<div onClick="javascript: getfile();" style="text-decoration : underline;cursor: pointer;">Download</div>
