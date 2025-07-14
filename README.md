# AutoLyrics Add-In for PowerPoint

A PowerPoint VSTO add-in that automatically imports song lyrics from text files and creates individual slides for each verse or chorus, perfect for worship services, karaoke presentations, or music performances.

## Why Choose This Free Alternative?

**No Monthly Fees** - Unlike subscription-based worship software that can cost $20-50+ per month, this add-in is completely free forever.

**Works with Your Existing Setup** - No need to learn new software or change your workflow. Uses PowerPoint that your team already knows.

**Perfect for Churches & Ministries** - Designed with budget-conscious organizations in mind:
- Small churches don't need expensive licensing
- No per-user fees for worship teams
- Youth groups and volunteer-run services can use it freely
- Mission organizations can deploy without ongoing costs

**Your Content, Your Control** - No cloud dependencies, account requirements, or online subscriptions. Your lyrics files stay on your computer.

**Flexible & Customizable** - Create your own slide layouts and styling instead of being locked into preset templates.

## Features

- **One-click import**: Select a text file containing song lyrics and automatically generate slides
- **Custom layout support**: Uses your existing "Lyrics" slide layout from the slide master
- **Smart text parsing**: Automatically separates verses/choruses based on paragraph breaks
- **Flexible text placement**: Intelligently finds content placeholders (excluding titles and footers)
- **Batch processing**: Creates multiple slides from a single lyrics file
- **Clean integration**: Adds seamlessly to PowerPoint's ribbon interface

## How It Works

1. Click the "Select Song" button in PowerPoint's ribbon
2. Choose a text file containing your lyrics
3. The add-in automatically:
   - Parses the lyrics into separate blocks (verses/choruses)
   - Creates a new slide for each block using your "Lyrics" layout
   - Populates the content placeholder with the lyrics text

## Requirements

- Microsoft PowerPoint (Office 365, 2019, 2021)
- .NET Framework 4.7.2 or later
- Visual Studio Tools for Office Runtime (VSTO Runtime)
- A custom slide layout named "Lyrics" in your presentation's slide master

## Installation

### Option 1: Simple Installation (Recommended)
1. Download the latest release files
2. Extract to a folder on your computer
3. Double-click the `.vsto` file to install
4. Follow the installation prompts

### Option 2: Manual Installation
1. Build the project in Visual Studio
2. Copy the output files to your desired location
3. Run the `.vsto` file to install

## Usage

### Setting Up Your Presentation
1. Create a custom slide layout called "Lyrics" in your slide master
2. Add a text placeholder for the lyrics content (don't check "Title" or "Footer")
3. Style the layout as desired for your lyrics display

### Importing Lyrics
1. Prepare your lyrics in a text file with verses/choruses separated by blank lines
2. Open PowerPoint and click "Select Song" in the ribbon
3. Choose your lyrics file
4. The add-in will create slides automatically

### Example Lyrics File Format
```
Verse 1 lyrics here
Multiple lines are supported

Chorus lyrics here
Also multiple lines

Verse 2 lyrics here
And so on...
```

## Development

Built with:
- Visual Basic.NET
- Visual Studio Tools for Office (VSTO)
- PowerPoint Object Model
- Windows Forms for file dialog

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues, questions, or feature requests, please open an issue on GitHub.

---

*Perfect for churches, music venues, karaoke nights, or any presentation where you need to display song lyrics professionally.*
