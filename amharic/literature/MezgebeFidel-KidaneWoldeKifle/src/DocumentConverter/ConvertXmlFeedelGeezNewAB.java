/*
 * Build: javac -cp icu4j-54_1_1.jar:commons-lang3-3.4.jar ConvertFeedelGeezNewAB.java 
 * Run:   java -cp icu4j-54_1_1.jar:commons-lang3-3.4.jar:. ConvertXmlFeedelGeezNewAB MezgebeFidel-KidaneWoldeKifle-GeezNewAB.xml 
 * 
 */
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
import java.nio.file.Files;
import static org.apache.commons.lang3.StringEscapeUtils.escapeXml10;
 

public class ConvertXmlFeedelGeezNewAB {
	  
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


	public void process(
		final String geezNewATextRulesFile,
		final String geezNewBTextRulesFile,
		final String inFile)
	{

		  try {
			// specify the transliteration file in the first argument.

			// read the input, transliterate, and write to output
			String geezNewATextFile = readRules( new File( geezNewATextRulesFile ) );
			String  geezNewBTextFile = readRules( new File(  geezNewBTextRulesFile ) );

			final Transliterator geezNewAText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin", geezNewATextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );
			final Transliterator  geezNewBText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin",  geezNewBTextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );

 
			// specify an outfile file in the 3rd argument.
			File header = new File( "word-xml-header.txt" );
			File body = new File( "outfile.xml" );
			body.delete();
			Files.copy( header.toPath(), body.toPath() );
			final BufferedWriter out = new BufferedWriter(
				 new OutputStreamWriter(
					 // new FileOutputStream( "outfile.xml" ), "UTF8" )
					 new FileOutputStream( body, true ), "UTF8" )
			);

			DefaultHandler handler = new DefaultHandler() {
				private String text = "";
				private String boundaryText = "";
				private boolean isGeezNewB = false;
				private boolean inDocument = false;
				Transliterator t = null;
				Attributes wtAttributes = null;

				private ArrayList<String> checks = new ArrayList<String>(
					Arrays.asList("H", "K", "c", "h", "m", "v", "z", "±", "Ñ", "Ù", "\u009E", "¤", "\u0085", "\u0099", "\u00ad", "\u00ae", "÷", "Ö", "ã", "W", "X", "ª", "ç", "ë", "ì" )

				);
				private ArrayList<String> selfClosing = new ArrayList<String>(
					Arrays.asList("w:sz", "w:rFonts" )

				);
				public boolean isIncomplete(String text) {
					if ( text.equals( "" ) ) { // || text.equals( "È" )) {
						return false;
					}
					return checks.contains( text.substring( text.length()-1 ) );
				}

				public String convertText( String text ) {
					String step1 = escapeXml10( t.transliterate( text ) );
					// String step2 = step1.replaceAll( "፡፡", "።"); // this usually won't work since each hulet neteb is surrounded by separate markup.
					String step2 = step1.replaceAll( "È", "»");
					return step2;
				}

				public void startElement(
					String namespaceURI,
				        String sName, // simple name
				        String qName, // qualified name
				        Attributes attrs)
					throws SAXException
				{
					if ( "w:document".equals( qName ) ) {
						inDocument = true;
					}
					if (! inDocument ) {
						return;
					}
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
							if( !text.equals( "" ) ) {
								boundaryText = convertText( text );
								text = "";
							}
							isGeezNewB = true;
						}
						writeOpenElement( qName, attrs, "w:cs", "Abyssinica SIL" );
					}
					else {
						// Print the starting tag verbatim
						if ( "w:t".equals( qName ) ) {
							wtAttributes = attrs;	
						}
						else {
							writeOpenElement( qName, attrs, null, null );
						}

					}
				}

				private Charset iso88591charset = Charset.forName("ISO-8859-1");
				private Charset utf8charset = Charset.forName("UTF-8");

				public void characters(
					char ch[],
					int start,
					int length)
					throws SAXException
				{
					// text += new String(ch, start, length);
					/*
					System.out.println( "Char Start: " + start );
					System.out.println( "Char Length: " + length );
					byte b = (byte)(ch[start] & 0x00ff);
					int i = (int)(ch[start] & 0xff);
					char c = (char)(i);
					*/
					char c;
					if( ch[start] == '' ) {
						c = ' ';
					}
					else {
						c = (char)(ch[start] & 0xff);
					}
					if ( c != '\n' ) {
						text += c;
					}
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

				public void endElement(
					String uri,
					String localName,
					String qName)
					throws SAXException
				{
					if (! inDocument ) {
						return;
					}
					if ( "w:document".equals( qName ) ) {
						inDocument = false;
					}
					if ( "w:t".equals( qName ) ) {
						// do conversion at this point
						// possibly the diacritical mark is in the following w:t tag
						if (! isIncomplete( text ) ) {
							char firstChar = text.charAt( 0 );
							char lastChar = text.charAt( text.length()-1 );
							if( Character.isWhitespace( firstChar )
							    || Character.isWhitespace( lastChar ) ) {
								writeOpenElement( qName, wtAttributes, "xml:space", "preserve" );
							}
							else {
								writeOpenElement( qName, wtAttributes, null, null );
							}
							
							try {
								if ( isGeezNewB ) {
									t = geezNewBText;
									isGeezNewB = false;
								}
								else {
									t = geezNewAText;
								}
								if(! boundaryText.equals("") ) {
									out.write( boundaryText );
									boundaryText = "";
								}
								out.write( convertText( text ) );
								out.write( "</"+qName+">" );
								out.flush();
							}
				  			catch(IOException ex) {
								  System.out.println( ex );
				  			}
							text = "";
						}
					}
					else if (! selfClosing.contains( qName ) ) {
						try {
							out.write( "</"+qName+">" );
							out.flush();
						}
			  			catch(IOException ex) {
							System.out.println( ex );
			  			}
		  			}
 
				}


				public void writeOpenElement(
					String qName,
					Attributes attrs,
					String extraAttr,
					String extraAttrValue)
					throws SAXException
				{
					try {
						out.write( "<" + qName );
						if (attrs != null) {
							for (int i = 0; i < attrs.getLength(); i++) {
								String aName = attrs.getQName(i);
								String aValue = attrs.getValue(i);
								out.write( " " );
								if( "GeezNewA".equals(aValue)
								    || "GeezNewB".equals(aValue) ) {
									aValue = "Abyssinica SIL";
								}
								out.write( aName + "=\"" + aValue + "\"" );
    							}
							if ( extraAttr != null ) {
								out.write( " " + extraAttr + " =\"" + extraAttrValue + "\"" );
							}
						}
						if ( selfClosing.contains( qName ) ) {
							out.write( "/" );
						}
						out.write( ">" );
						out.flush();
					}
		  			catch(IOException ex) {
						System.out.println( ex );
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

			BufferedReader footer = new BufferedReader(new FileReader( "word-xml-footer.txt" ) );
			String line;
			while( (line = footer.readLine()) != null ) {
				out.write( line );
			}
			out.flush();
			out.close();
		  }
		  catch(Exception ex) {
		 	System.out.println( ex );
		  }
	}


	public static void main(String[] args) {
		ConvertXmlFeedelGeezNewAB t = new ConvertXmlFeedelGeezNewAB();
		t.process( "GeezNewATable.txt", "GeezNewBTable.txt", args[0] );
	}

}
