import unittest

import omnitag_mobile


class AppleModelResolutionTests(unittest.TestCase):
    def test_resolves_known_ipad_from_pymobiledevice_catalog(self):
        self.assertEqual(
            omnitag_mobile.resolver_nombre_apple("iPad13,4", "iPad"),
            "iPad Pro 11-inch (3rd gen, WiFi)",
        )

    def test_unknown_ipad_is_not_reported_as_iphone(self):
        self.assertEqual(
            omnitag_mobile.resolver_nombre_apple("iPad99,1", "iPad"),
            "iPad Desconocido",
        )

    def test_manual_model_is_saved_for_current_device(self):
        app = object.__new__(omnitag_mobile.OmniTagMobileApp)
        app.current_device_info = {"udid": "ipad-test"}
        app._manual_model_overrides = {}

        app._guardar_modelo_manual("iPad Pro 13 M4 1TB")

        self.assertEqual(
            app._manual_model_overrides["ipad-test"],
            "iPad Pro 13 M4 1TB",
        )


if __name__ == "__main__":
    unittest.main()
