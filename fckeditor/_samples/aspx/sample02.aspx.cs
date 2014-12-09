using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;

namespace FredCK.FCKeditorV2.Samples
{
	public class Sample02 : System.Web.UI.Page
	{
		protected System.Web.UI.HtmlControls.HtmlTextArea txtSubmitted;
		protected System.Web.UI.HtmlControls.HtmlInputButton btnSubmit;
		protected System.Web.UI.HtmlControls.HtmlGenericControl eSubmittedDataBlock;

		protected FredCK.FCKeditorV2.FCKeditor FCKeditor1;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// We'll do initializartion settings in the editor only the first time.
			// Once the page is submitted, nothing must be done in the OnLoad.
			if ( Page.IsPostBack )
				return ;

			eSubmittedDataBlock.Visible = false ;

			// Automatically calculates the editor base path based on the _samples directory.
			// This is usefull only for these samples. A real application should use something like this:
			// FCKeditor1.BasePath = '/FCKeditor/' ;	// '/FCKeditor/' is the default value.
			string sPath = Request.Url.AbsolutePath ;
			int iIndex = sPath.LastIndexOf( "_samples") ;
			sPath = sPath.Remove( iIndex, sPath.Length - iIndex  ) ;
			
			FCKeditor1.BasePath = sPath ;
			FCKeditor1.Value = "This is some <strong>sample text</strong>. You are using <a href=\"http://www.fckeditor.net/\">FCKeditor</a>." ;
			FCKeditor1.ImageBrowserURL	= sPath + "editor/filemanager/browser/default/browser.html?Type=Image&Connector=connectors/aspx/connector.aspx" ;
			FCKeditor1.LinkBrowserURL	= sPath + "editor/filemanager/browser/default/browser.html?Connector=connectors/aspx/connector.aspx" ;
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: This call is required by the ASP.NET Web Form Designer.
			//
			InitializeComponent();
			base.OnInit(e);

			// This events declarations has been moved to the "OnInit" method
			// to avoid a bug in Visual Studio to delete then without any advice.
			this.Load += new System.EventHandler(this.Page_Load);
			this.btnSubmit.ServerClick += new System.EventHandler(this.btnSubmit_ServerClick);
		}
		
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{    

		}
		#endregion

		private void btnSubmit_ServerClick(object sender, System.EventArgs e)
		{
			eSubmittedDataBlock.Visible = true ;
			txtSubmitted.Value = HttpUtility.HtmlEncode( FCKeditor1.Value ) ;
		}
	}
}
