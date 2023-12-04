using Microsoft.WindowsAPICodePack.Net;
using Spectre.Console;
using System.Collections.Generic;
using System.Threading;
using System;
using System.Linq;

public static class Program
{
	static Dictionary<string, Network> networkMap = new Dictionary<string, Network>();
	static bool exit = false;
	static Style highlightStyle = new Style().Foreground(Color.Aqua);
	public static void Main(string[] args)
	{
		AnsiConsole.Status().Start("Starting Up", ctx =>
		{
			if (System.Diagnostics.Debugger.IsAttached)
				Thread.Sleep(250); // Helps start up faster, believe it or not.
		});

		while (!exit)
			Stage1();
	}
	private static void Stage1()
	{
		AnsiConsole.Clear();
		networkMap.Clear();

		Table table = new Table()
			.LeftAligned()
			.Border(TableBorder.Rounded);

		table.AddColumn("[aqua]Network[/]");
		table.AddColumn("[aqua]Category[/]");

		NetworkCollection networks = NetworkListManager.GetNetworks(NetworkConnectivityLevels.Connected);

		// Iterate over all connected networks
		foreach (Network network in networks)
		{
			string name = GetNetworkName(network);
			int counter = 2;
			while (networkMap.ContainsKey(name))
			{
				name = GetNetworkName(network) + " (" + counter + ")";
				counter++;
			}
			networkMap[name] = network;
		}

		foreach (KeyValuePair<string, Network> pair in networkMap.OrderBy(k => k.Key))
		{
			string name = pair.Key;
			Network network = pair.Value;

			string category = GetNetworkCategoryText(network.Category);

			table.AddRow("[aqua]" + name + "[/]", category);
		}

		AnsiConsole.Write(table);

		List<string> choices = new List<string>(new[]{
				"Set all networks to Public.", // 0
				"Set all networks to Private.", // 1
				"Set all networks to Domain.", // 2
				"Set some networks to Public.", // 3
				"Set some networks to Private.", // 4
				"Set some networks to Domain.", // 5
				"Refresh Network List", // 6 
				"Exit (CTRL + C)" // 7
		});
		SelectionPrompt<string> selectionPrompt = new SelectionPrompt<string>();
		selectionPrompt.Title("Choose an option.");
		selectionPrompt.AddChoices(choices);
		selectionPrompt.HighlightStyle(highlightStyle);

		string choice = selectionPrompt.Show(AnsiConsole.Console);
		switch (choices.IndexOf(choice))
		{
			case 0:
				SetDomainsAll(NetworkCategory.Public);
				break;
			case 1:
				SetDomainsAll(NetworkCategory.Private);
				break;
			case 2:
				SetDomainsAll(NetworkCategory.Authenticated);
				break;
			case 3:
				SetDomainsPrompt(NetworkCategory.Public);
				break;
			case 4:
				SetDomainsPrompt(NetworkCategory.Private);
				break;
			case 5:
				SetDomainsPrompt(NetworkCategory.Authenticated);
				break;
			case 6:
				return;
			case 7:
				exit = true;
				return;
		}
	}

	private static void SetDomainsPrompt(NetworkCategory category)
	{
		string catText = GetNetworkCategoryText(category);

		AnsiConsole.Clear();

		string[] options = networkMap.Where(k => k.Value.Category != category)
			.Select(k => k.Key)
			.OrderBy(k => k)
			.ToArray();

		if (options.Length == 0)
		{
			AnsiConsole.Foreground = Color.Yellow;
			AnsiConsole.WriteLine("All networks are already in category " + catText);
			AnsiConsole.WriteLine();
			AnsiConsole.ResetColors();
			AnsiConsole.WriteLine("Press any key to restart...");
			Console.ReadKey();
			return;
		}

		MultiSelectionPrompt<string> multiSelectionPrompt = new MultiSelectionPrompt<string>();
		multiSelectionPrompt.Title("Select networks to set to " + catText);
		multiSelectionPrompt.AddChoices(options);
		multiSelectionPrompt.HighlightStyle(highlightStyle);

		if (options.Length > 10)
		{
			multiSelectionPrompt
				.PageSize(3)
				.MoreChoicesText("(Move up and down to reveal more options)");
		}

		List<string> choices = multiSelectionPrompt.Show(AnsiConsole.Console);

		List<string> errors = new List<string>();
		foreach (string choice in choices)
		{
			if (networkMap.TryGetValue(choice, out Network n))
			{
				try
				{
					n.Category = category;
				}
				catch (Exception ex)
				{
					errors.Add("Failed to set network category of network \"" + choice + "\" to " + catText + ": " + ex.Message);
				}
			}
		}
		if (errors.Count > 0)
		{
			AnsiConsole.Foreground = Color.Red;
			AnsiConsole.WriteLine();
			foreach (string error in errors)
				AnsiConsole.WriteLine(error);
			AnsiConsole.WriteLine();
			AnsiConsole.ResetColors();
			AnsiConsole.WriteLine("Press any key to restart...");
			Console.ReadKey();
		}
	}

	private static void SetDomainsAll(NetworkCategory category)
	{
		string catText = GetNetworkCategoryText(category);
		List<string> errors = new List<string>();

		NetworkCollection networks = NetworkListManager.GetNetworks(NetworkConnectivityLevels.Connected);

		// Iterate over all connected networks
		foreach (Network network in networks)
		{
			try
			{
				network.Category = category;
			}
			catch (Exception ex)
			{
				errors.Add("Failed to set network category of network \"" + network.Name + "\" to " + catText + ": " + ex.Message);
			}
		}

		if (errors.Count > 0)
		{
			AnsiConsole.Foreground = Color.Red;
			AnsiConsole.WriteLine();
			foreach (string error in errors)
				AnsiConsole.WriteLine(error);
			AnsiConsole.WriteLine();
			AnsiConsole.ResetColors();
			AnsiConsole.WriteLine("Press any key to restart...");
			Console.ReadKey();
		}
	}

	private static string GetNetworkCategoryText(NetworkCategory ncat)
	{
		string category = ncat.ToString();
		if (ncat == NetworkCategory.Authenticated)
			category = "Domain";
		return category;
	}
	private static string GetNetworkName(Network network)
	{
		string name = network.Name;
		if (network.IsConnectedToInternet)
			name += " (internet)";
		return name;
	}
}