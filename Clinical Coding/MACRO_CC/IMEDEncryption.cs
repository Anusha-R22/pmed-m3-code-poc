using System;
using XceedEncryptionLib;
using XceedZipLib;
using System.Text;

namespace InferMed.MACRO.ClinicalCoding.MACRO_CC
{
	/// <summary>
	/// Summary description for csEncryption.
	/// </summary>
	public class IMEDEncryption
	{
		private const string _EXCEED_ZIP_LICENCE = "SFX45-XHUMC-NTHKJ-G45A";
		private const string _EXCEED_BIN_ENCODE_LICENCE = "BEN10-5RUZC-TTYTN-AAXA";
		private const string _EXCEED_ENCRYPTION_LICENCE = "CRY10-ARUXC-ATY7N-B4JA";
		private const string _SECRETKEY = "Zorba the Greek";

		private IMEDEncryption()
		{
			//
			// TODO: Add constructor logic here
			//
		}

		public static string DecryptString(string sString)
		{
			XceedEncryptionLib.XceedEncryptionClass xEncryptor = new XceedEncryptionLib.XceedEncryptionClass();
			xEncryptor.License( _EXCEED_ENCRYPTION_LICENCE );
			XceedZipLib.XceedCompressionClass xCompress = new XceedZipLib.XceedCompressionClass();
			XceedEncryptionLib.XceedRijndaelEncryptionMethodClass xRijndael = new XceedEncryptionLib.XceedRijndaelEncryptionMethodClass();
			object oString = new object();
			object oDecrypted = new object();
			object oUncompressedData = new object();
			XceedZipLib.xcdCompressionError xResult = new XceedZipLib.xcdCompressionError();
			string sDecrypted = "";

			try
			{
				xCompress.License( _EXCEED_ZIP_LICENCE );

				//set the encryption parameters
				xRijndael.EncryptionMode = 0;
				xRijndael.PaddingMethod = 0;

				//set the secret key
				xRijndael.SetSecretKeyFromPassPhrase( _SECRETKEY, 256 );

				//set the encryption method
				xEncryptor.EncryptionMethod = xRijndael;

				//decrypt the string
				oString = ConvertToObjectByteArray( sString );
				oDecrypted = xEncryptor.Decrypt( ref oString, true );

				//uncompresses the decrypted data
				xResult = xCompress.Uncompress( ref oDecrypted, out oUncompressedData, true );
				if ( xResult == XceedZipLib.xcdCompressionError.xceSuccess )
				{
					//convert the ascii byte array to string
					sDecrypted = ConvertToString( oUncompressedData );
				}

				return( sDecrypted );
			}
			finally
			{
				
			}
		}
		// encrypt connection string
		public static string EncryptString(string sString)
		{
			string sEncrypted = "";
			XceedEncryptionLib.XceedEncryptionClass xEncryptor = new XceedEncryptionLib.XceedEncryptionClass();
			xEncryptor.License( _EXCEED_ENCRYPTION_LICENCE );
			XceedEncryptionLib.XceedRijndaelEncryptionMethodClass xRijndael = new XceedEncryptionLib.XceedRijndaelEncryptionMethodClass();
			XceedZipLib.XceedCompressionClass xCompress = new XceedZipLib.XceedCompressionClass();
			XceedZipLib.xcdCompressionError xResult = new XceedZipLib.xcdCompressionError();
			object oCompressedData = new object();

			try
			{
				// set license keys
				xCompress.License( _EXCEED_ZIP_LICENCE );

				// set encryption method params
				xRijndael.EncryptionMode = 0;
				xRijndael.PaddingMethod = 0;

				// set the secret key
				xRijndael.SetSecretKeyFromPassPhrase( _SECRETKEY, 256 );

				// set the encryption method
				xEncryptor.EncryptionMethod = xRijndael;

				object oDecryptedText = ConvertToObjectANSIByteArray( sString );

				// compress the data before encryption
				xResult = xCompress.Compress( ref oDecryptedText, out oCompressedData, true );
			    
				if(xResult == xcdCompressionError.xceSuccess)
				{
					//encrypt the string and convert it from binary to hex
					sEncrypted = BinaryToHex( ( byte[] )xEncryptor.Encrypt( ref oCompressedData, true ) );
				}
				else
				{
					//raise error ? - xCompress.GetErrorDescription(xResult)
				}

				return sEncrypted;
			}
			finally
			{
			}
		}

		// convert string to an ANSI byte array
		private static object ConvertToObjectANSIByteArray(string sString)
		{
			// get ANSI byte string
			byte[] byteString = Encoding.Default.GetBytes( sString );

			return ( object )byteString;
		}

		// convert binary array to a hex encoded string
		private static string BinaryToHex(byte[] byteBinary)
		{
			StringBuilder sbHex = new System.Text.StringBuilder();
			string sHex;
			for( int i=0; i < byteBinary.Length; i++ )
			{
				sHex="0" + System.Convert.ToString(byteBinary[i],16);
				sHex=(sHex.Length==2)? sHex.Substring(0,2): sHex.Substring(1,2);
				sbHex.Append(sHex);
			}
			return sbHex.ToString();
		}
		private static object ConvertToObjectByteArray(string sData)
		{
			byte[] b = new byte[( sData.Length / 2 )];

			for( int i = 0; i < sData.Length; i += 2 )
			{
				b[i / 2] = System.Convert.ToByte( sData.Substring( i, 2 ), 16 );
			}
			return( ( object )b );
		}

		private static string ConvertToString(object oData)
		{
			System.Text.Encoding unicode = System.Text.Encoding.Unicode;
			System.Text.Encoding ascii = System.Text.Encoding.ASCII;

			//convert the decrypted object into a string
			byte[] bData = (byte[])oData;
			byte[] bUnicodeBytes = System.Text.Encoding.Convert(ascii,unicode,bData);

			char[] cUnicodeChars = new char[unicode.GetCharCount(bUnicodeBytes, 0, bUnicodeBytes.Length)];
			unicode.GetChars(bUnicodeBytes, 0, bUnicodeBytes.Length, cUnicodeChars, 0);
			string sUnicodeString = new string(cUnicodeChars);

			return sUnicodeString;
		}
	}
}
