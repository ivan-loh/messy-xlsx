"""Tests for shared handler utility functions."""

from io import BytesIO
from pathlib import Path

from messy_xlsx.parsing.base_handler import (
    get_file_desc,
    is_fileobj,
    read_file_content,
    reset_buffer,
)


class TestIsFileobj:
    def test_bytesio_is_fileobj(self) -> None:
        assert is_fileobj(BytesIO()) is True

    def test_path_is_not_fileobj(self) -> None:
        assert is_fileobj(Path("test.xlsx")) is False


class TestResetBuffer:
    def test_resets_position(self) -> None:
        buf = BytesIO(b"hello")
        buf.read(3)
        assert buf.tell() == 3
        reset_buffer(buf)
        assert buf.tell() == 0

    def test_no_seek_does_not_raise(self) -> None:
        class NoSeek:
            pass

        reset_buffer(NoSeek())  # should not raise


class TestGetFileDesc:
    def test_stream_description(self) -> None:
        assert get_file_desc(BytesIO()) == "<stream>"

    def test_path_description(self) -> None:
        p = Path("/tmp/test.xlsx")
        assert "test.xlsx" in get_file_desc(p)


class TestReadFileContent:
    def test_reads_all_content(self) -> None:
        buf = BytesIO(b"hello world")
        buf.read(5)  # advance position
        result = read_file_content(buf)
        assert result == b"hello world"
