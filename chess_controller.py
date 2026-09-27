"""Draw a python-chess board in an open WPS Spreadsheet workbook."""

import os
from pathlib import Path
from typing import Any

import chess
import chess.engine
import win32com.client


PIECE_UNICODE = {
	"P": "♙",
	"N": "♘",
	"B": "♗",
	"R": "♖",
	"Q": "♕",
	"K": "♔",
	"p": "♟",
	"n": "♞",
	"b": "♝",
	"r": "♜",
	"q": "♛",
	"k": "♚",
	".": "",
}


def connect_to_wps() -> Any:
	"""Return the active WPS application, or start one when none is active."""
	try:
		return win32com.client.GetActiveObject("Ket.Application")
	except (AttributeError, OSError):
		try:
			return win32com.client.GetActiveObject("ET.Application")
		except (AttributeError, OSError):
			return win32com.client.Dispatch("Ket.Application")


def draw_board(
	board: chess.Board,
	sheet: Any,
	start_row: int = 3,
	start_column: int = 3,
) -> None:
	"""Write the board's 8x8 position to WPS, starting at C3 by default."""
	for row in range(8):
		for column in range(8):
			square = chess.square(column, 7 - row)
			piece = board.piece_at(square)
			symbol = piece.symbol() if piece else "."
			sheet.Cells(start_row + row, start_column + column).Value = (
				PIECE_UNICODE[symbol]
			)


def create_stockfish_engine(
	stockfish_path: str,
	skill_level: int | None = 5,
) -> chess.engine.SimpleEngine:
	"""Start Stockfish and optionally apply its Skill Level setting."""
	engine = chess.engine.SimpleEngine.popen_uci(stockfish_path)
	if skill_level is not None:
		if not 0 <= skill_level <= 20:
			engine.quit()
			raise ValueError("skill_level must be between 0 and 20")
		engine.configure({"Skill Level": skill_level})
	return engine


def get_stockfish_path() -> str:
	"""Use STOCKFISH_PATH or stockfish.exe beside this script."""
	environment_path = os.environ.get("STOCKFISH_PATH")
	path = Path(environment_path) if environment_path else Path(__file__).with_name(
		"stockfish.exe"
	)
	if not path.is_file():
		raise FileNotFoundError(f"Stockfish not found: {path}")
	return str(path)


def main() -> None:
	"""Connect WPS, draw the initial position, and start Stockfish."""
	stockfish_path = get_stockfish_path()

	wps = connect_to_wps()
	sheet = wps.ActiveWorkbook.ActiveSheet
	board = chess.Board()
	draw_board(board, sheet)

	engine = create_stockfish_engine(stockfish_path)
	try:
		print(engine.analyse(board, chess.engine.Limit(depth=12)))
	finally:
		engine.quit()


if __name__ == "__main__":
	main()
