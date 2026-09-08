import unittest
import urllib.parse
import zipfile
from pathlib import Path

from scraper import (
    build_rows,
    chunks_by_product,
    fetch_products,
    load_config,
    OutputRow,
    resolve_attribute,
    select_temu_category,
    temu_color,
    temu_size_selection,
    variation_size,
)


def product(name, *slugs):
    return {"name": name, "categories": [{"slug": slug} for slug in slugs]}


class ScraperLogicTests(unittest.TestCase):
    def test_null_attribute_is_empty_instead_of_literal_none(self):
        self.assertEqual(resolve_attribute({}, "Размер", None), "")

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

    def test_temu_regular_alpha_size_fields(self):
        self.assertEqual(
            temu_size_selection("XXXL"),
            ("2 - Regular Size", "10 - Alpha", "3XL"),
        )

    def test_temu_bra_size_uses_custom_family(self):
        self.assertEqual(
            temu_size_selection("65F"),
            ("101 - Custom size", "10 - Alpha", "65F"),
        )

    def test_temu_one_size_fields(self):
        self.assertEqual(
            temu_size_selection("One Size"),
            ("2 - Regular Size", "1 - One Size", "one-size"),
        )

    def test_size_chart_image_is_added_to_every_output_row(self):
        source = {
            "id": 1,
            "name": "Тестов сутиен",
            "categories": [{"slug": "sutieni"}],
            "prices": {"price": "2000", "regular_price": "2500", "currency_minor_unit": 2},
            "images": [{"src": "https://example.com/product.jpg"}],
            "is_in_stock": True,
        }
        config = load_config(Path("/definitely/missing/config.json"))
        output = build_rows([source], config)[0].values
        self.assertEqual(output["t_5_Size Chart Method"], "Upload size chart image")
        self.assertEqual(
            output["t_5_Size Chart Image"],
            "https://valea.bg/wp-content/uploads/2025/04/size-table.jpg",
        )

    def test_missing_variation_colors_are_inferred_from_declared_terms(self):
        source = {
            "id": 2,
            "name": "Тестов продукт",
            "categories": [{"slug": "bikini"}],
            "attributes": [
                {"name": "Размер", "has_variations": True, "terms": [{"name": "M", "slug": "m"}]},
                {"name": "Цвят", "has_variations": True, "terms": [
                    {"name": "Бял", "slug": "byal"}, {"name": "Черно", "slug": "cherno"}
                ]},
            ],
            "variations": [
                {"id": 21, "attributes": [{"name": "Размер", "value": "m"}, {"name": "Цвят", "value": None}]},
                {"id": 22, "attributes": [{"name": "Размер", "value": "m"}, {"name": "Цвят", "value": None}]},
            ],
            "prices": {"price": "2000", "regular_price": "2500", "currency_minor_unit": 2},
            "images": [],
        }
        stats = {}
        rows = build_rows([source], load_config(Path("/missing")), stats)
        self.assertEqual([row.color for row in rows], ["White", "Black"])
        self.assertEqual(stats["inferred_color_variations"], 2)

    def test_invalid_and_duplicate_variations_are_skipped(self):
        source = {
            "id": 3,
            "name": "Тестов продукт",
            "categories": [{"slug": "bikini"}],
            "attributes": [
                {"name": "Размер", "has_variations": True, "terms": [{"name": "L", "slug": "l"}]},
                {"name": "Цвят", "has_variations": True, "terms": [{"name": "Кафяво", "slug": "kafyavo"}]},
            ],
            "variations": [
                {"id": 31, "attributes": [{"name": "Размер", "value": "l"}, {"name": "Цвят", "value": None}]},
                {"id": 32, "attributes": [{"name": "Размер", "value": "l"}, {"name": "Цвят", "value": None}]},
                {"id": 33, "attributes": [{"name": "Размер", "value": None}, {"name": "Цвят", "value": None}]},
            ],
            "prices": {"price": "2000", "regular_price": "2500", "currency_minor_unit": 2},
            "images": [],
        }
        stats = {}
        rows = build_rows([source], load_config(Path("/missing")), stats)
        self.assertEqual(len(rows), 1)
        self.assertEqual(stats["skipped_duplicate_variations"], 1)
        self.assertEqual(stats["skipped_invalid_variations"], 1)

    def test_chunks_keep_parent_product_together(self):
        def row(product_id):
            return OutputRow(product_id, str(product_id), "1", "M", "Black", {})

        rows = [row(1), row(1), row(1), row(2), row(2), row(2)]
        parts = list(chunks_by_product(rows, 4))
        self.assertEqual([len(part) for part in parts], [3, 3])
        self.assertEqual([{item.product_id for item in part} for part in parts], [{1}, {2}])

    def test_generated_workbook_is_a_valid_zip_archive(self):
        from tempfile import TemporaryDirectory
        from scraper import TemuWorkbook

        source = {
            "id": 4,
            "name": "Тестов сутиен",
            "categories": [{"slug": "sutieni"}],
            "prices": {"price": "2000", "regular_price": "2500", "currency_minor_unit": 2},
            "images": [{"src": "https://example.com/product.jpg"}],
            "is_in_stock": True,
        }
        rows = build_rows([source], load_config(Path("/missing")))
        template = Path(__file__).parents[1] / "template" / "TEMU_TEMPLATE.xlsx"
        with TemporaryDirectory() as directory:
            output = Path(directory) / "output.xlsx"
            TemuWorkbook(template).write(output, rows)
            with zipfile.ZipFile(output) as archive:
                self.assertIsNone(archive.testzip())
            self.assertFalse(output.with_name(output.name + ".tmp").exists())

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
