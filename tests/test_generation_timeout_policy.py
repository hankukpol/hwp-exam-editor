import unittest

from core.models import ExamDocument
from gui.main_window import GenerationWorker


class GenerationTimeoutPolicyTestCase(unittest.TestCase):
    @staticmethod
    def _worker(question_count: int, file_type: str = "TYPE_A", style_required: bool = True) -> GenerationWorker:
        doc = ExamDocument(file_type=file_type, subject="", questions=[], total_count=question_count)
        return GenerationWorker(
            document=doc,
            source_file="dummy.hwp",
            timeout_sec=60,
            style_required=style_required,
        )

    def test_timeout_scales_up_for_large_question_count(self) -> None:
        w20 = self._worker(20, "TYPE_A", True)
        w100 = self._worker(100, "TYPE_A", True)
        self.assertGreater(w100.timeout_sec, w20.timeout_sec)
        self.assertGreater(w100.timeout_sec, 60)

    def test_retry_budgets_scale_for_large_question_count(self) -> None:
        worker = self._worker(100, "TYPE_A", True)
        self.assertGreaterEqual(worker._compute_rpc_retry_timeout_sec(), 120)
        self.assertGreaterEqual(worker._compute_safe_retry_timeout_sec(), 120)
        self.assertGreater(worker.guard_timeout_sec(), worker.timeout_sec)


if __name__ == "__main__":
    unittest.main()
