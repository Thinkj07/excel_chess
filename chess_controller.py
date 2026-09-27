"""Draw a python-chess board in an open WPS Spreadsheet workbook."""

import os
import time
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
			cell = sheet.Cells(start_row + row, start_column + column)
			cell.Value = PIECE_UNICODE[symbol]
			cell.Font.Color = 255 if piece and piece.color == chess.WHITE else 0


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


def set_cell_value(sheet: Any, cell: str, value: str) -> None:
	"""Write a value to a WPS cell."""
	row, column = cell_coordinates(cell)
	sheet.Cells(row, column).Value = value


def get_cell_value(sheet: Any, cell: str) -> str:
	"""Read a WPS cell as trimmed text."""
	row, column = cell_coordinates(cell)
	value = sheet.Cells(row, column).Value
	return str(value).strip() if value is not None else ""


def cell_coordinates(cell: str) -> tuple[int, int]:
	"""Convert an A1-style cell address into WPS row and column numbers."""
	letters = "".join(character for character in cell.upper() if character.isalpha())
	digits = "".join(character for character in cell if character.isdigit())
	if not letters or not digits:
		raise ValueError(f"Invalid cell address: {cell}")

	column = 0
	for character in letters:
		column = column * 26 + ord(character) - ord("A") + 1
	return int(digits), column


def game_over_message(board: chess.Board) -> str | None:
	"""Return the result message when the game has ended."""
	if board.is_checkmate():
		winner = "Trắng" if board.turn == chess.BLACK else "Đen"
		return f"Chiếu bí! {winner} thắng."
	if board.is_stalemate():
		return "Ván cờ hòa do pat."
	if board.is_repetition(3):
		return "Ván cờ hòa do lặp nước."
	if board.is_insufficient_material():
		return "Ván cờ hòa do không đủ quân chiếu bí."
	return None


def play_game(
	sheet: Any,
	engine: chess.engine.SimpleEngine,
	poll_interval: float = 0.5,
) -> None:
	"""Run a WPS polling loop with the human playing White."""
	board = chess.Board()
	draw_board(board, sheet)
	set_cell_value(sheet, "L3", "Nhập nước đi của bạn và nhấn Enter.")
	set_cell_value(sheet, "N3", "")
	set_cell_value(sheet, "L4", "")
	set_cell_value(sheet, "L6", "Đến lượt bạn (Trắng).")
	set_cell_value(sheet, "L8", "")
	set_cell_value(sheet, "L9", "Nước đi của Stockfish")

	while True:
		if get_cell_value(sheet, "N3").casefold() == "reset":
			board = chess.Board()
			set_cell_value(sheet, "N3", "")
			set_cell_value(sheet, "L4", "")
			set_cell_value(sheet, "L8", "")
			draw_board(board, sheet)
			set_cell_value(sheet, "L6", "Đến lượt bạn (Trắng).")
			continue

		result_message = game_over_message(board)
		if result_message:
			set_cell_value(sheet, "L6", result_message)
			return

		if board.turn == chess.WHITE:
			move_text = get_cell_value(sheet, "L4")
			if not move_text:
				time.sleep(poll_interval)
				continue
			try:
				move = chess.Move.from_uci(move_text.lower())
			except ValueError:
				set_cell_value(sheet, "L6", "Nước đi sai luật, vui lòng nhập lại")
				time.sleep(poll_interval)
				continue
			if move not in board.legal_moves:
				set_cell_value(sheet, "L6", "Nước đi sai luật, vui lòng nhập lại")
				time.sleep(poll_interval)
				continue

			board.push(move)
			set_cell_value(sheet, "L4", "")
			draw_board(board, sheet)
			continue

		set_cell_value(sheet, "L6", "Stockfish đang tính...")
		result = engine.play(board, chess.engine.Limit(time=1.0))
		if result.move is None:
			set_cell_value(sheet, "L6", "Không tìm thấy nước đi hợp lệ.")
			return
		board.push(result.move)
		set_cell_value(sheet, "L8", result.move.uci())
		draw_board(board, sheet)
		set_cell_value(sheet, "L6", "Đến lượt bạn (Trắng).")


def main() -> None:
	"""Connect WPS and run the chess game loop."""
	stockfish_path = get_stockfish_path()

	wps = connect_to_wps()
	sheet = wps.ActiveWorkbook.ActiveSheet

	engine = create_stockfish_engine(stockfish_path)
	try:
		play_game(sheet, engine)
	finally:
		engine.quit()


if __name__ == "__main__":
	main()
