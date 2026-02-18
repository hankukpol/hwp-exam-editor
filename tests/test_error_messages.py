import unittest

from core.error_messages import build_generation_error_message, build_parse_error_message


class ErrorMessageTestCase(unittest.TestCase):
    def test_generation_rpc_message(self) -> None:
        message = build_generation_error_message(
            "(-2147023174, 'RPC 서버를 사용할 수 없습니다.', None, None)"
        )
        self.assertIn("연결", message)
        self.assertIn("Hwp.exe", message)
        self.assertIn("[원본 오류]", message)

    def test_generation_default_message(self) -> None:
        message = build_generation_error_message("unknown failure")
        self.assertIn("출력 생성", message)
        self.assertIn("원본 오류", message)

    def test_parse_encrypted_message(self) -> None:
        message = build_parse_error_message("파일이 암호로 보호되어 있습니다.")
        self.assertIn("암호", message)
        self.assertIn("원본 오류", message)

    def test_parse_question_number_message_is_preserved(self) -> None:
        raw = "문제 번호를 찾지 못했습니다.\n인식 점검: ..."
        self.assertEqual(build_parse_error_message(raw), raw)


if __name__ == "__main__":
    unittest.main()

