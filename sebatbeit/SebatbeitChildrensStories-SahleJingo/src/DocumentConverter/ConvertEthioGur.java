import java.io.*;
import com.ibm.icu.text.*;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;
 

public class ConvertEthioGur {
	  
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


	public void process(final String plainTextRulesFile, final String boldTextRulesFile, final String inFile) {

		  try {
			// specify the transliteration file in the first argument.

			// read the input, transliterate, and write to output
			String plainTextFile = readRules( new File( plainTextRulesFile ) );
			String  boldTextFile = readRules( new File(  boldTextRulesFile ) );

			final Transliterator plainText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin", plainTextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );
			final Transliterator  boldText = Transliterator.createFromRules( "Ethiopic-ExtendedLatin",  boldTextFile.replace( '\ufeff', ' ' ), Transliterator.REVERSE );

 
			// specify an outfile file in the 3rd argument.
			final BufferedWriter out = new BufferedWriter(
				 new OutputStreamWriter(
					 new FileOutputStream( "outfile.txt" ), "UTF8" )
			 );

			DefaultHandler handler = new DefaultHandler() {
				private String text = null;
				private boolean bold = false;
				Transliterator t = null;

				public void startElement(String namespaceURI,
				        String sName, // simple name
				        String qName, // qualified name
				        Attributes attrs)
				throws SAXException
				{
					if ( "w:rFonts".equals( qName ) ) {
						String typeface = attrs.getValue( "w:ascii" );
						if( "ETHIOGUR".equals( typeface ) ) {
							System.out.println( "Found: ETHIOGUR" );
						}
					}
					if ( "w:t".equals( qName ) ) {
						System.out.println( "Text to Convert" );
						text = "";
						
					}
					if ( "w:b".equals( qName ) || "w.bCs".equals( qName ) ) {
						System.out.println( "  BOLD ON" );
						bold = true;
					}
					return;
				}
 
				public void characters(char ch[], int start, int length) throws SAXException {
					text += new String(ch, start, length);
				}

				public void endElement(String uri, String localName, String qName) throws SAXException {
					if ( "w:t".equals( qName ) ) {
						System.out.println( "Text Off" );
						// do conversion at this point
						try {
							if (bold ) {
								t = boldText;
								bold = false;
							}
							else {
								t = plainText;
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
		ConvertEthioGur t = new ConvertEthioGur();
		t.process( "EthioGurage-PlainText.txt", "EthioGurage-BoldText.txt", args[0] );
	}

}
