import unittest

from adbeam_excel_parser.icp_autosearch import (
    FetchedHtml,
    build_acronym_variants,
    build_direct_domain_urls,
    build_direct_domain_variants,
    build_lab_variants,
    extract_contact_data,
    score_domain_zone,
    score_search_candidate,
    unwrap_google_url,
)


class IcpAutoSearchTests(unittest.TestCase):
    def test_direct_domain_urls_include_ampersand_without_and(self) -> None:
        urls = build_direct_domain_urls("Art&Fact")

        self.assertIn("https://artfact.ru/", urls)
        self.assertIn("https://art-fact-products.com/", urls)

    def test_direct_domain_variants_remove_descriptor_words(self) -> None:
        self.assertIn("aravia", build_direct_domain_variants("Aravia Professional"))
        self.assertIn("letique", build_direct_domain_variants("Letique Cosmetics"))

    def test_direct_domain_variants_include_acronyms_and_lab_split(self) -> None:
        self.assertIn("dtmskin", build_direct_domain_variants("Don't Touch My Skin"))
        self.assertIn("esti-lab", build_direct_domain_variants("Estilab"))
        self.assertIn("dtmskin", build_acronym_variants(["don", "t", "touch", "my", "skin"]))
        self.assertIn("esti-lab", build_lab_variants(["estilab"]))

    def test_direct_domain_variants_include_transliterated_cyrillic_shortcuts(self) -> None:
        self.assertIn("https://kpcosm.ru/", build_direct_domain_urls("Краснополянская косметика"))
        self.assertIn("https://master-om.com/", build_direct_domain_urls("Мастерская Олеси Мустаевой"))
        self.assertIn("https://trives-spb.ru/", build_direct_domain_urls("Trives"))
        self.assertIn("https://venoshop.ru/", build_direct_domain_urls("Venoteks"))
        self.assertIn("https://theblackpearl.ru/", build_direct_domain_urls("Черный Жемчуг"))
        self.assertIn("https://estel.beauty/", build_direct_domain_urls("Юникосметик"))
        self.assertIn("https://chistayalinia.ru/", build_direct_domain_urls("Чистая Линия"))

    def test_domain_zone_prefers_ru_for_russian_base(self) -> None:
        self.assertGreater(score_domain_zone("example.ru"), score_domain_zone("example.fr"))

    def test_unwrap_google_url_extracts_target(self) -> None:
        self.assertEqual(
            unwrap_google_url("/url?q=https%3A%2F%2Fkpcosm.ru%2F&sa=U"),
            "https://kpcosm.ru/",
        )

    def test_google_style_official_result_scores_cyrillic_brand(self) -> None:
        score = score_search_candidate(
            brand="Краснополянская косметика",
            segment="КОСМЕТИКА, ПАРФЮМЕРИЯ, ГИГИЕНА",
            url="https://kpcosm.ru/",
            title="Краснополянская косметика — официальный сайт",
            snippet="Бренд уходовых продуктов натурального происхождения.",
        )

        self.assertGreaterEqual(score, 45)

    def test_multiword_brand_rejects_generic_partial_match(self) -> None:
        self.assertLess(
            score_search_candidate(
                brand="Doctor Wax Russia",
                segment="АВТОТОВАРЫ, МАСЛА, ИНСТРУМЕНТ DIY",
                url="https://doctor.ru/",
                title="Доктор.ру - портал о здоровье",
                snippet=None,
            ),
            0,
        )
        self.assertLess(
            score_search_candidate(
                brand="Магнит / магнитные доски Iqaktiv",
                segment="БЫТОВАЯ ТЕХНИКА И ЭЛЕКТРОНИКА",
                url="https://magnit.ru/",
                title="Доставка продуктов на дом - Магнит",
                snippet="Купить готовую еду и продукты в Магнит",
            ),
            0,
        )
        self.assertLess(
            score_search_candidate(
                brand="Magnit",
                segment="БЫТОВАЯ ТЕХНИКА И ЭЛЕКТРОНИКА",
                url="https://www.magnit.com/ru/",
                title="Magnit",
                snippet=None,
            ),
            0,
        )
        self.assertLess(
            score_search_candidate(
                brand="Magnit",
                segment="БЫТОВАЯ ТЕХНИКА И ЭЛЕКТРОНИКА",
                url="https://magnit.ru/",
                title="Доставка продуктов на дом - Магнит",
                snippet="Купить готовую еду и продукты в Магнит",
            ),
            0,
        )

    def test_cosmetics_segment_penalizes_wrong_game_candidate(self) -> None:
        score = score_search_candidate(
            brand="Art&Fact",
            segment="КОСМЕТИКА, ПАРФЮМЕРИЯ, ГИГИЕНА",
            url="https://artandfact.ru/",
            title="Art&Fact - студия разработки игр",
            snippet=None,
        )

        self.assertLess(score, 35)

    def test_extract_contact_data_filters_test_email(self) -> None:
        data = extract_contact_data([
            FetchedHtml(
                url="https://example.ru/contacts/",
                status_code=200,
                html="""
                <html><body>
                <a href="mailto:test@example.ru">test@example.ru</a>
                <a href="mailto:email@example.ru">email@example.ru</a>
                <a href="mailto:feb084cf563d213c704f083c073ecbb3@o4505635341205504.ingest.sentry.io">sentry</a>
                <a href="mailto:hello@realbrand.ru">hello@realbrand.ru</a>
                Телефон +7 (800) 123-45-67. ИНН 1234567890.
                <a href="https://t.me/example">Telegram</a>
                </body></html>
                """,
            )
        ])

        self.assertEqual(data["emails"], ["hello@realbrand.ru"])
        self.assertEqual(data["phones"], ["+7 800 123-45-67"])
        self.assertEqual(data["inns"], ["1234567890"])
        self.assertEqual(data["social_links"], ["https://t.me/example"])


if __name__ == "__main__":
    unittest.main()
