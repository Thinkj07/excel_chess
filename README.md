# Chess Controller for WPS

Control a chess game between a player and Stockfish through WPS Spreadsheets.
The board is drawn in cells `C3` through `J10`.

## Requirements

- Windows
- WPS Spreadsheets
- Python 3.10 or later
- `stockfish.exe` in the same folder as `chess_controller.py`

A different Stockfish path can be provided through the `STOCKFISH_PATH` environment variable.

## Installation

Open PowerShell in the project folder and run:

```powershell
python -m pip install python-chess pywin32
```

## Running the Program

Open a workbook in WPS, select the sheet to use, and then run:

```powershell
python chess_controller.py
```

The program connects to the active workbook and draws the initial position.

## Usage

| Cell     | Function                           |
| -------- | ---------------------------------- |
| `C3:J10` | Chess board                        |
| `L3`     | Move entry instructions            |
| `L4`     | UCI move input, for example `e2e4` |
| `L6`     | Game status, errors, or result     |
| `L8`     | Stockfish's most recent move       |
| `N3`     | Enter `Reset` to restart the game  |

The player uses the White pieces, which are displayed in red. Stockfish uses the Black pieces, which are displayed in black.

After entering a move in `L4`, press Enter. Moves must use UCI notation, for example:

```text
e2e4
```

If the move is legal, the board is updated and Stockfish calculates a response for up to one second.

## Notes

- WPS must be running with a workbook open.
- The `STOCKFISH_PATH` environment variable takes priority over the `stockfish.exe` file in the project folder.
