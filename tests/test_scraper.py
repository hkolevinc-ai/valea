import unittest
import urllib.parse

from scraper import fetch_products, select_temu_category, temu_color, variation_size


def product(name, *slugs):
    return {"name": name, "categories": [{"slug": slug} for slug in slugs]}


class ScraperLogicTests(unittest.TestCase):
    def test_category_mapping(self):
        cases = [
            (product("Дамски сутиен Jennifer Soft", "sutieni"), "29094"),
            (product("Сутиен за кърмачки", "sutieni", "za-karmachki"), "29091"),
            (product("Дамски бикини Rosetta Figi Big", "bikini"), "29081"),
            (product("Дамски бикини Paula Figi", "bikini"), "29082"),
            (product("Дамски боксерки", "bikini", "bokserki"), "29083"),
            (product("Дамски прашки", "prashki"), "29086"),
            (product("Дамско боди", "body"), "29110"),
            (product("Дамски корсет", "korseti"), "29103"),
            (product("Дамски колан", "kolani"), "29120"),
            (product("Дамски топ", "top"), "53993"),
            (product("Еротичен комплект", "erotichno-belyo"), "29096"),
            (product("Дамски бебидол", "erotichno-belyo"), "29105"),
        ]
        for source, expected in cases:
            with self.subTest(source=source):
                self.assertEqual(select_temu_category(source), expected)

    def test_bra_size_is_combined_in_temu_order(self):
        self.assertEqual(
            variation_size([("Чашка размер", "F"), ("Подгръдна обиколка", "65")]),
            "65F",
        )

    def test_color_mapping(self):
        self.assertEqual(temu_color(["Черно"]), "Black")
        self.assertEqual(temu_color(["Черно", "Червено"]), "Multicolor")

    def test_api_page_size_falls_back_after_server_error(self):
        class FakeClient:
            def __init__(self):
                self.page_sizes = []

            def get_json(self, url):
                query = urllib.parse.parse_qs(urllib.parse.urlsplit(url).query)
                page_size = int(query["per_page"][0])
                page = int(query["page"][0])
                self.page_sizes.append(page_size)
                if page_size > 10:
                    raise RuntimeError("HTTP Error 500: Internal Server Error")
                items = [
                    {"id": 101, "parent": 0, "categories": [{"slug": "sutieni"}]},
                    {"id": 102, "parent": 0, "categories": [{"slug": "bikini"}]},
                ]
                return (items if page == 1 else []), {"x-wp-totalpages": "1"}

        client = FakeClient()
        result = fetch_products(client)
        self.assertEqual([item["id"] for item in result], [101, 102])
        self.assertEqual(client.page_sizes, [25, 10])


if __name__ == "__main__":
    unittest.main()
