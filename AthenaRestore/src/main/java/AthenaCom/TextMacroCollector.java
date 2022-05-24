package AthenaCom;

import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URL;

public class TextMacroCollector {

    private final String MACRO_URL = "https://api.preview.platform.athenahealth.com/v1/{6545}/textmacros";

    public TextMacroCollector () {

    }

    public void GET () {
        try {
            final URL url = new URL( MACRO_URL );
            try {
                final HttpURLConnection conn = (HttpURLConnection) url.openConnection();
                conn.setRequestMethod( "GET" );
                conn.connect();
                final int responseCode = conn.getResponseCode();

                if ( responseCode != 200 ) {
                    throw new RuntimeException( "HttpResponseCode: " + responseCode );
                }
                else {

                }

            }
            catch ( final IOException e ) {
                e.printStackTrace();
            }
        }
        catch ( final MalformedURLException e ) {
            e.printStackTrace();
        }
    }
}
