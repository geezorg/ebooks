import java.io.*;
import com.ibm.icu.text.*;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
import java.nio.charset.Charset;
import java.nio.ByteBuffer;
import java.nio.CharBuffer;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.ArrayList;
 

public class ConvertFeedelGeezNewAB {
	  
	public String readRules( File fileName ) throws IOException {
		  String line, segment, rules = "";
		  BufferedReader ruleFile = new BufferedReader(  new InputStreamReader(new FileInputStream( fileName.toString() ), "UTF-8"));
		  while ( (line = ruleFile.readLine()) != null) {
			  if ( line.trim().equals("") || line.charAt(0) == '#' ) {
				  continue;
			  }
			  segment = line.replaceFirst ( "^(.*?)#(.*)$", "$1" );
			  rules += ( segment == null ) ? line : segment;
		  }
		  
		  return rules;
	}


	public void process(final String geezNewATextRulesFile, final String geezNewBTextRulesFile, final String inFile) {

		  try {
			// specify the transliteration file in the first argument.

			// read the input, transliterate, and write to output
			String geezNewATextFile = readRules( new File( geezNewATextRulesFile ) );
			String  geezNewBTextFile = readRules( new File(  geezNewBTextRulesFile ) );

			final Transliterator geezNewAText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin", geezNewATextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );
			final Transliterator  geezNewBText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin",  geezNewBTextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );

 
			// specify an outfile file in the 3rd argument.
			final BufferedWriter out = new BufferedWriter(
				 new OutputStreamWriter(
					 new FileOutputStream( "outfile.txt" ), "UTF8" )
			 );

			DefaultHandler handler = new DefaultHandler() {
				private String text = null;
				private boolean isGeezNewB = false;
				Transliterator t = null;

				private ArrayList<String> checks = new ArrayList<String>(
					Arrays.asList("H", "K", "c", "h", "m", "v", "z", "±", "Ñ", "Ù", "\u009E", "¤", "\u0085", "\u0099", "\u00ad", "\u00ae", "÷", "Ö", "ã", "W", "X", "ª", "ç", "ë", "ì") //, "\"")

				);
				public boolean isIncomplete(String text) {
					if ( (text == null) || text.equals( "" ) || text.equals( "È" )) {
						return false;
					}
					// String lastChar = String.valueOf( text.charAt( text.length()-1 ) );
					String lastChar = text.substring( text.length()-1 );
					return checks.contains( lastChar );
				}

				public void startElement(String namespaceURI,
				        String sName, // simple name
				        String qName, // qualified name
				        Attributes attrs)
				throws SAXException
				{
					if ( "w:rFonts".equals( qName ) ) {
						String typeface = attrs.getValue( "w:ascii" );
						if( "GeezNewA".equals( typeface ) ) {
							// System.out.println( "Found: GeezNewA" );
							isGeezNewB = false;
						}
						else if( "GeezNewB".equals( typeface ) ) {
							// We need this check in the event that an
							// "incomplete" GeezNewA character is still
							// in the buffer at the time a change to GeezNewB
							// occurs.  This is common for text like: ከ፮፻፺፬ 
							// Otherwise 0x9e for ከ will not be transliterated.
							// So here we check if there is anyting in the buffer
							// and convert before flushing and filling the buffer
							// with GeezNewB letters.
							if( ((text != null) && !text.equals( "" )) ) {
								try {
									out.write( t.transliterate( text ) );
									text = "";
								}
			  					catch(IOException ex) {
									  System.out.println( ex );
					  			}
							}
							// System.out.println( "Found: GeezNewB" );
							isGeezNewB = true;
						}
					}
					if ( "w:t".equals( qName ) ) {
						// System.out.println( "Text to Convert" );
						if (! isIncomplete( text ) ) {
							text = "";
						}
						
					}
					return;
				}

				private Charset iso88591charset = Charset.forName("ISO-8859-1");
				private Charset utf8charset = Charset.forName("UTF-8");
				public void characters(char ch[], int start, int length) throws SAXException {
					// text += new String(ch, start, length);
					/*
					System.out.println( "Char Start: " + start );
					System.out.println( "Char Length: " + length );
					byte b = (byte)(ch[start] & 0x00ff);
					int i = (int)(ch[start] & 0xff);
					char c = (char)(i);
					*/
					char c = (char)(ch[start] & 0xff);
					text += c;
					/*
					System.out.println( "Char: " + ch[start] );
					System.out.println( "Code: " + (int)ch[start] );
					System.out.println( "b: " + b );
					System.out.println( "c: " + c );
					System.out.println( "int: " + i );
					System.out.println( "Text: " + text );
					*/
/*
					System.out.println( "Length: " + text.length() );
					byte[] bytes = text.getBytes(StandardCharsets.UTF_16); 
					// byte bytes[] = text.getBytes( iso88591charset );
					System.out.println( "Byte[0]: " + (int)bytes[0] );
					System.out.println( "Byte[1]: " + (int)bytes[1] );
					String string2 = new String ( bytes, iso88591charset );
					System.out.println( "Converted: " + string2 );

// decode UTF-8
ByteBuffer encoded = iso88591charset.encode(text);
CharBuffer cbuf = iso88591charset.decode( encoded );
// String data = new String( cbuf, StandardCharsets.ISO_8859_1 );

// encode ISO-8559-1
// ByteBuffer outputBuffer = iso88591charset.decode(inputBuffer);
// byte[] outputData = outputBuffer.array();

System.out.println( "Decoded: " + cbuf );
*/
				}

				public void endElement(String uri, String localName, String qName) throws SAXException {
					if ( "w:t".equals( qName ) ) {
						// System.out.println( "Text Off" );
						// do conversion at this point
						if ( isIncomplete( text ) ) {
							return;
						}
						try {
							if ( isGeezNewB ) {
								t = geezNewBText;
								isGeezNewB = false;
							}
							else {
								t = geezNewAText;
							}
							if( text.startsWith( "\\p" ) ) {
								out.write( "\n<p>\n" );
								text = text.substring(2);
							}
							else if( text.startsWith( "\\s" ) ) {
								out.write( "\n<h2></h2>\n" );
								text = text.substring(2);
							}
							out.write( t.transliterate( text ) );
						}
			  			catch(IOException ex) {
							  System.out.println( ex );
			  			}
						text = null;
					}
					else if ( "w:p".equals( qName ) ) {
						try {
							  out.write( "\n" );
			  			}
			  			catch(IOException ex) {
							  System.out.println( ex );
			  			}
					}
 
				}
 
     			};
 
			SAXParserFactory factory = SAXParserFactory.newInstance();
			SAXParser saxParser = factory.newSAXParser();

			File file = new File( inFile );
			InputStream inputStream= new FileInputStream(file);
			Reader reader = new InputStreamReader(inputStream,"UTF-8");
 
			InputSource is = new InputSource(reader);
			is.setEncoding("UTF-8");

		  	// InputStreamReader is = new InputStreamReader(new FileInputStream( inFile ), "UTF-8");
 
			saxParser.parse(is, handler);

			out.flush();
			out.close();
		  }
		  catch(Exception ex) {
		 	System.out.println( ex );
		  }
	}


	public static void main(String[] args) {
		ConvertFeedelGeezNewAB t = new ConvertFeedelGeezNewAB();
		t.process( "GeezNewATable.txt", "GeezNewBTable.txt", args[0] );
	}

}
