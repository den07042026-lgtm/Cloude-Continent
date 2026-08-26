import sys
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import ozon_stock_sync as sync


class UpdateOzonStocksTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        sync.log.disabled = True

    @classmethod
    def tearDownClass(cls):
        sync.log.disabled = False

    def test_retries_transient_network_error(self):
        response = Mock(status_code=200)
        response.raise_for_status.return_value = None
        response.json.return_value = {"result": [{"offer_id": "a-con", "errors": []}]}

        with patch.object(sync.requests, "post", side_effect=[OSError("temporary"), response]) as post, \
             patch.object(sync.time, "sleep"):
            updated, errors = sync.update_ozon_stocks("client", "key", 1, {"a-con": 4})

        self.assertEqual((updated, errors), (1, 0))
        self.assertEqual(post.call_count, 2)

    def test_reports_failed_batch_after_four_attempts(self):
        with patch.object(sync.requests, "post", side_effect=OSError("offline")) as post, \
             patch.object(sync.time, "sleep"):
            updated, errors = sync.update_ozon_stocks("client", "key", 1, {"a-con": 4})

        self.assertEqual((updated, errors), (0, 1))
        self.assertEqual(post.call_count, 4)

    def test_rejects_partial_success_response(self):
        response = Mock(status_code=200)
        response.raise_for_status.return_value = None
        response.json.return_value = {"result": []}

        with patch.object(sync.requests, "post", return_value=response) as post, \
             patch.object(sync.time, "sleep"):
            updated, errors = sync.update_ozon_stocks("client", "key", 1, {"a-con": 4})

        self.assertEqual((updated, errors), (0, 1))
        self.assertEqual(post.call_count, 4)


if __name__ == "__main__":
    unittest.main()
