using Spectre.Console;

namespace GrabConf;

public static class Log
{
    public static bool Verbose { get; set; }

    public static void Info(string message)
        => AnsiConsole.MarkupLine($"[blue][[info]][/] {message.EscapeMarkup()}");

    public static void Success(string message)
        => AnsiConsole.MarkupLine($"[green][[done]][/] {message.EscapeMarkup()}");

    public static void Error(string message)
        => AnsiConsole.MarkupLine($"[red][[error]][/] {message.EscapeMarkup()}");

    public static void Detail(string message)
        => AnsiConsole.MarkupLine($"  [dim]{message.EscapeMarkup()}[/]");

    public static void Debug(string message)
    {
        if (Verbose)
            AnsiConsole.MarkupLine($"[grey][[debug]] {message.EscapeMarkup()}[/]");
    }
}
