using System;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization.Formatters.Soap;
using System.IO;
using System.Text;

namespace Lapointe.SharePoint.PowerShell.StsAdm.OperationHelpers
{
	/// <summary>
	/// Serialization helper methods.  These methods must be used to ensure consistent serializing and
	/// deserializing of data (especially when versioning is concerned).
	/// </summary>
	public abstract class SerializationHelper
	{

		/// <summary>
		///     Default constructor.  Is private so that no one can instantiate an instance of the class.
		/// </summary>
		/// 
		/// <returns>
		///     A void Serialization value...
		/// </returns>
        private SerializationHelper()
		{
		}

		/// <summary>
		///     Serialize an object to Xml.  Requires that the object supports serialization and can in fact be
		///     serialized to Xml.  Will throw an exception if it cannot be converted.
		/// </summary>
		/// <param name="o" type="object">
		///     <para>
		///         The object to serialize.
		///     </para>
		/// </param>
		/// <returns>
		///     Xml text representing the object or empty string if object is null
		/// </returns>
		public static string XmlSerializationHelper(object o)
		{
			return XmlSerializationHelper(o, false);
		}

		/// <summary>
		///     Serialize an object to Xml.  Requires that the object supports serialization and can in fact be
		///     serialized to Xml.  Will throw an exception if it cannot be converted.
		/// </summary>
		/// <param name="o">The object to serialize.</param>
		/// <param name="includeXmlDeclarations">if set to <c>true</c> will include XML declarations in result.</param>
		/// <returns>Xml text representing the object or empty string if object is null</returns>
		public static string XmlSerializationHelper(object o, bool includeXmlDeclarations)
		{
			if (o == null)
				return "";

			StringWriter output = new StringWriter(new StringBuilder());
			string ret;

			try
			{
				XmlSerializer s = new XmlSerializer(o.GetType());
				s.Serialize(output, o);

				// To cut down on the size of the xml being sent to the database, we'll strip
				// out this extraneous xml.
				ret = output.ToString();
				if (!includeXmlDeclarations)
				{
					ret = ret.Replace("xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"", "");
					ret = ret.Replace("xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"", "");
					ret = ret.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "").Trim();
				}
			}
			catch (Exception ex)
			{
                Exception ex1 = new Exception("Unable to serialize object!", ex);
				throw ex1;
			}

			return ret;
		}

		/// <summary>
		///     Deserialize an xml string to its object graph.
		/// </summary>
		/// <param name="s" type="string">
		///     <para>
		///         The serialized string to deserialize.
		///     </para>
		/// </param>
		/// <param name="type" type="System.Type">
		///     <para>
		///         The resultant object type.
		///     </para>
		/// </param>
		/// <returns>
		///     The deserialized object.
		/// </returns>
		public static object XmlDeSerializationHelper(string s, Type type)
		{
			// Create an instance of the XmlSerializer specifying type and namespace.
			XmlSerializer serializer = new XmlSerializer(type);

			MemoryStream stream = new MemoryStream(UTF8Encoding.UTF8.GetBytes(s));

			XmlReader reader = new XmlTextReader(stream);
          
			object o;
			try
			{
				o = serializer.Deserialize(reader);
			}
			catch (Exception ex) 
			{
                Exception ex1 = new Exception("Unable to deserialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}
			return o;

		}

		/// <summary>
		///     Serialize an object to a base64 string.  Requires that the object supports serialization.
		///     Will throw an exception if it cannot be converted.
		/// </summary>
		/// <param name="o" type="object">
		///     <para>
		///         The object to serialize.
		///     </para>
		/// </param>
		/// <returns>
		///     A Base64 encoded string.
		/// </returns>
		public static string BinarySerializationHelper(object o)
		{
			byte[] data;
			BinarySerializationHelper(o, out data);
			return Convert.ToBase64String(data);
		}

		/// <summary>
		///     Serialize an object to a byte array.  Requires that the object supports serialization.
		///     Will throw an exception if it cannot be converted.
		/// </summary>
		/// <param name="o" type="object">
		///     <para>
		///         
		///     </para>
		/// </param>
		/// <param name="data" type="byte[]">
		///     <para>
		///         
		///     </para>
		/// </param>
		/// <returns>
		///     A void value...
		/// </returns>
		public static void BinarySerializationHelper(object o, out byte[] data)
		{
			MemoryStream stream = new MemoryStream();
			try
			{
				BinaryFormatter binaryFormatter = new BinaryFormatter();
				binaryFormatter.AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple;
				binaryFormatter.Serialize(stream, o);

				data = stream.ToArray();
			}
			catch (Exception ex) 
			{  
				Exception ex1 = new Exception("Unable to serialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}
		}

		/// <summary>
		///     Deserialize an object from a base64 binary serialized string.
		/// </summary>
		/// <param name="s" type="string">
		///     <para>
		///         The string to deserialize.
		///     </para>
		/// </param>
		/// <returns>
		///     The deserialized object.
		/// </returns>
		public static object BinaryDeSerializationHelper(string s)
		{
			MemoryStream stream;
			stream = new MemoryStream(Convert.FromBase64String(s));
			object o;
			try
			{
				BinaryFormatter binaryFormatter = new BinaryFormatter();
				binaryFormatter.AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple;
				o = binaryFormatter.Deserialize(stream);
			}
			catch (Exception ex) 
			{  
				Exception ex1 = new Exception("Unable to deserialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}

			return o;
		}

		/// <summary>
		///     Deserialize an object from a byte array.
		/// </summary>
		/// <param name="data" type="byte[]">
		///     <para>
		///         The string to deserialize.
		///     </para>
		/// </param>
		/// <returns>
		///     The deserialized object.
		/// </returns>
		public static object BinaryDeSerializationHelper(byte[] data)
		{
			MemoryStream stream;
			stream = new MemoryStream(data);
			object o;
			try
			{
				BinaryFormatter binaryFormatter = new BinaryFormatter();
				binaryFormatter.AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple;
				o = binaryFormatter.Deserialize(stream);
			}
			catch (Exception ex) 
			{  
				Exception ex1 = new Exception("Unable to deserialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}

			return o;
		}

		/// <summary>
		///     Serialize an object to a base64 or plain text string.  Requires that the object supports serialization.
		///     Will throw an exception if it cannot be converted.
		/// </summary>
		/// <param name="o" type="object">
		///     <para>
		///         The object to serialize.
		///     </para>
		/// </param>
		/// <param name="base64Encode" type="bool">
		///     <para>
		///         True: convert the output to a base64 string (use this if storing to a database or file or transmitting)
		///         False: convert the output to a human readable xml string (use this if you need to read or manipulate the output)
		///     </para>
		/// </param>
		/// <returns>
		///     The serialized object.
		/// </returns>
		public static string SoapSerializationHelper(object o, bool base64Encode)
		{
			string s;
			MemoryStream stream = new MemoryStream();
			try
			{
				SoapFormatter soapFormatter = new SoapFormatter();
				soapFormatter.AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple;
				soapFormatter.Serialize(stream, o);

				if (base64Encode)
					s = Convert.ToBase64String(stream.ToArray());
				else
					s = UTF8Encoding.UTF8.GetString(stream.ToArray());
			}
			catch (Exception ex) 
			{
                Exception ex1 = new Exception("Unable to serialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}
			return s;
		}

		/// <summary>
		///     Deserialize an object from a base64 or plain text serialized string.
		/// </summary>
		/// <param name="s" type="string">
		///     <para>
		///         The string to deserialize.
		///     </para>
		/// </param>
		/// <param name="base64Encoded" type="bool">
		///     <para>
		///         True: the string was base64 encoded.
		///         False: the string is plain text xml.
		///     </para>
		/// </param>
		/// <returns>
		///     The deserialized object.
		/// </returns>
		public static object SoapDeSerializationHelper(string s, bool base64Encoded)
		{
			MemoryStream stream;
			object o;
			if (base64Encoded)
				stream = new MemoryStream(Convert.FromBase64String(s));
			else
				stream = new MemoryStream(UTF8Encoding.UTF8.GetBytes(s));

			try
			{
				SoapFormatter soapFormatter = new SoapFormatter();
				soapFormatter.AssemblyFormat = System.Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple;
				o = soapFormatter.Deserialize(stream);
			}
			catch (Exception ex) 
			{
                Exception ex1 = new Exception("Unable to deserialize object!", ex);
				throw ex1; 
			}
			finally
			{
				stream.Close();
			}

			return o;
		}

	}
}
