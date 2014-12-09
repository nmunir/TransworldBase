/*
 * FCKeditor - The text editor for internet
 * Copyright (C) 2003-2004 Frederico Caldeira Knabben
 * 
 * Licensed under the terms of the GNU Lesser General Public License:
 * 		http://www.opensource.org/licenses/lgpl-license.php
 * 
 * For further information visit:
 * 		http://www.fckeditor.net/
 * 
 * File Name: Util.cs
 * 	Useful tools.
 * 
 * Version:  2.1
 * Modified: 2005-02-19 16:58:19
 * 
 * File Authors:
 * 		Frederico Caldeira Knabben (fredck@fckeditor.net)
 */

using System ;
using System.Runtime.InteropServices ;
using System.IO ;
using System.Collections ;

namespace FredCK.FCKeditorV2
{
	public sealed class Util
	{
		// The "_mkdir" function is used by the "CreateDirectory" method.
		[DllImport("msvcrt.dll", SetLastError=true)]
		private static extern int _mkdir(string path) ;

		private Util()
		{}

		/// <summary>
		/// This method should provide safe substitude for Directory.CreateDirectory().
		/// </summary>
		/// <param name="path">The directory path to be created.</param>
		/// <returns>A <see cref="System.IO.DirectoryInfo"/> object for the created directory.</returns>
		/// <remarks>
		///		<para>
		///		This method creates all the directory structure if needed.
		///		</para>
		///		<para>
		///		The System.IO.Directory.CreateDirectory() method has a bug that gives an
		///		error when trying to create a directory and the user has no rights defined
		///		in one of its parent directories. The CreateDirectory() should be a good 
		///		replacement to solve this problem.
		///		</para>
		/// </remarks>
		public static DirectoryInfo CreateDirectory( string path )
		{
			ArrayList oDirsToCreate = new ArrayList() ;

			// Create the directory info object for that dir (normalized to its absolute representation).
			DirectoryInfo oDir = new DirectoryInfo( Path.GetFullPath( path ) ) ;

			// Check the entire path structure to find directories that must be created.
			while ( oDir != null && !oDir.Exists )
			{
				oDirsToCreate.Add( oDir.FullName ) ;
				oDir = oDir.Parent ;
			}

			// "oDir == null" means that the check arrives in the root and it doesn't exist too.
			if ( oDir == null )
				throw( new System.IO.DirectoryNotFoundException( "Directory \"" + oDirsToCreate[ oDirsToCreate.Count-1 ] + "\" not found." ) ) ;

			// Create all directories that must be created (from bottom to top).
			for( int i = oDirsToCreate.Count - 1 ; i >= 0 ; i-- )
			{
				string sPath = (string)oDirsToCreate[i] ;
				int iReturn = _mkdir( sPath ) ;

				if ( iReturn != 0 )
					throw new ApplicationException("Error calling [msvcrt.dll]:_wmkdir(" + sPath + "), error code: " + iReturn );
			}

			return new DirectoryInfo( path ) ;
		}
	}
}
