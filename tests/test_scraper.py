import unittest

from scraper import select_temu_category, temu_color, variation_size


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


if __name__ == "__main__":
    unittest.main()
