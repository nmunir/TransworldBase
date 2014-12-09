            function clickButton(e, buttonid){
              var evt = e ? e : window.event;
              var bt = document.getElementById(buttonid);
              if (bt){
                  if (evt.keyCode == 13){
                        bt.click();
                        return false;
                      }
                  }
            }

           function TT_OpenHelpWindow(value){
                window.open(value,"Help","top=10,left=10,width=500,height=400,status=no,toolbar=no,address=no,menubar=no,resizable=no,scrollbars=yes");
           }


           function TT_PrintConsignment(Key){
                window.open("ConsignmentNote.aspx?key=" + Key,"Consignment","top=10,left=10,width=625,height=550,status=no,toolbar=yes,address=no,menubar=no,resizable=yes,scrollbars=no");
           }

           function SB_ShowImage(value){
                window.open("show_image.aspx?Image=" + value,"ProductImage","top=10,left=10,width=610,height=610,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=yes");
           }

           function CMShowUsage(value){
                window.open("CMShowUsage.aspx?ref=" + value,"Usage","top=10,left=10,width=610,height=610,status=no,toolbar=no,address=no,menubar=no,resizable=yes,scrollbars=yes");
           }
