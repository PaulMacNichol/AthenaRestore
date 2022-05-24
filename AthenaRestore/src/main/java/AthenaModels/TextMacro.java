package AthenaModels;

public class TextMacro {

    private String shortcut;
    private String users;
    private String section;
    private String expansion;

    public TextMacro () {

    }

    public TextMacro ( final String shortcut, final String users, final String section, final String expansion ) {
        setShortcut( correctMacroName( shortcut ) );
        setUsers( users );
        setSection( section );
        setExpansion( expansion );
    }

    private String correctMacroName ( final String macroName ) {
        final StringBuilder sb = new StringBuilder( macroName );

        for ( int i = 0; i < macroName.length(); i++ ) {
            if ( i == 0 ) {
                sb.setCharAt( 0, Character.toUpperCase( sb.charAt( 0 ) ) );
            }
            if ( sb.charAt( i ) == '_' && i < sb.length() - 1 ) {
                sb.setCharAt( i + 1, Character.toUpperCase( sb.charAt( i + 1 ) ) );
            }

            if ( sb.charAt( i ) == '\\' || sb.charAt( i ) == '/' || sb.charAt( i ) == ' ' ) {
                sb.setCharAt( i, '_' );
            }
        }

        return sb.toString();
    }

    /**
     * @return the shortcut
     */
    public String getShortcut () {
        return shortcut;
    }

    /**
     * @param shortcut
     *            the shortcut to set
     */
    public void setShortcut ( final String shortcut ) {
        this.shortcut = shortcut;
    }

    /**
     * @return the users
     */
    public String getUsers () {
        return users;
    }

    /**
     * @param users
     *            the users to set
     */
    public void setUsers ( final String users ) {
        this.users = users;
    }

    /**
     * @return the section
     */
    public String getSection () {
        return section;
    }

    /**
     * @param section
     *            the section to set
     */
    public void setSection ( final String section ) {
        this.section = section;
    }

    /**
     * @return the expansion
     */
    public String getExpansion () {
        return expansion;
    }

    /**
     * @param expansion
     *            the expansion to set
     */
    public void setExpansion ( final String expansion ) {
        this.expansion = expansion;
    }

    @Override
    public String toString () {
        return "TextMacro [shortcut=" + shortcut + ", users=" + users + ", section=" + section + ", expansion="
                + expansion + "]";
    }
}
