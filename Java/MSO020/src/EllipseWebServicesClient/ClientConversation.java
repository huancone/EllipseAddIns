package EllipseWebServicesClient;

public class ClientConversation
{
    // Fields
    public static String password;
    public static String username;

    // Methods
    public static void authenticate(String username, String password)
    {
       ClientConversation.username = username;
       ClientConversation.password = password;
    }
    
}
