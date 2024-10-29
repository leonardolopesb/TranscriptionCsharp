using System;

public class Account
{
    public string Email { get; private set; }
    public string Password { get; private set; }

    public Account (string email, string password)
    {
        Email = email;
        Password = password;
    }

    public static Account LoadData(string txt)
    {
        string[] lines = File.ReadAllLines(txt);

        string email = lines[0];
        string password = lines[1];

        return new Account (email, password);
    }
}