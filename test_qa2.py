import os
import tempfile
import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, patch

import QA2


class AuditWorkspaceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        cls.application = QA2.QApplication.instance() or QA2.QApplication([])

    def setUp(self):
        self.window = QA2.AuditReviewDemo()
        section = self.window.add_section("Structure 1: Main device")
        for style, status in (("6575C88G16", "PASS"), ("42C2066G31", "UNRESOLVED")):
            item = QA2.QTreeWidgetItem(section, [
                style, "Test assembly", "Cover", "TOPCOVER", "Width = 30", "", "", "", "",
                "G16", "G16" if status == "PASS" else "", status,
            ])
            item.setToolTip(11, "Validation evidence")
            item.setCheckState(0, QA2.Qt.CheckState.Unchecked)
            self.window.checkbox_items[style] = item
        self.window.update_progress()

    def tearDown(self):
        self.window.close()

    def test_search_and_status_filters_preserve_audit_rows(self):
        self.window.status_filter.setCurrentText("PASS")
        self.assertEqual(self.window.visible_rows.text(), "1 of 2 assemblies")
        self.window.audit_search.setText("42C2066")
        self.assertEqual(self.window.visible_rows.text(), "0 of 2 assemblies")
        self.window.status_filter.setCurrentText("All statuses")
        self.assertEqual(self.window.visible_rows.text(), "1 of 2 assemblies")
        self.assertEqual(len(self.window.checkbox_items), 2)
        self.assertEqual(self.window.summary_labels["PASS"].text(), "1")

    def test_navigation_column_toggle_and_details(self):
        self.window.navigation_group.button(1).click()
        self.assertEqual(self.window.main_tabs.currentIndex(), 1)
        self.assertEqual(self.window.bom_list.header().visualIndex(10), 3)
        self.assertTrue(self.window.bom_list.isColumnHidden(4))
        self.window.parameter_toggle.setChecked(True)
        self.assertFalse(self.window.bom_list.isColumnHidden(4))
        self.window._show_row_details(self.window.checkbox_items["6575C88G16"])
        self.assertIn("Validation evidence", self.window.row_details.toPlainText())
        self.assertIn("Width = 30", self.window.row_details.toPlainText())

    def test_section_requirement_is_used_in_display_and_comparison(self):
        window = self.window
        window._code_mappings = QA2.parse_code_mappings((
            ("Name", "Code", "Parameter 1", "Parameter 2", "Parameter 3"),
            ("Bus", "WCCTAUX3", "Amperage", "Material", "Heat Test/Density"),
        ))
        window._assembly_families["2C11531G03"] = "WCCTAUX3"
        window._assembly_links["2C11531G03"] = "WCCTAUX3.xlsx"
        window.parameter_variables = QA2.ExcelParameterVariables.from_sources([
            "Bus Amps: 3000\nBus Material: Copper\nHeat Test",
        ])
        window._section_bus_requirements = QA2.section_bus_requirements(
            "SECTION 8 NEEDS DENSITY RATED BUS (998P160)", "outline.dwg",
        )
        with patch.object(window._excel_com, "get_group_tables", return_value=()), \
                patch.object(window._excel_com, "get_sheet_values", return_value=WorkbookValueTests.wcctaux3_sheet()):
            for number in (6, 8):
                section = window.add_section(f"Structure {number}")
                window.add_checkbox(section, f"STRUCT{number}::2C11531G03", "2C11531G03", "XBUS,UGPS,25-3000,CU DEN,51W", "51", "36", "51", "36", number)
        item = window.checkbox_items["STRUCT8::2C11531G03"]
        self.assertEqual([item.text(column) for column in (9, 10, 11)], ["G03", "G03", "PASS"])
        self.assertIn("Heat Test/Density = DENSITY", item.text(6))
        self.assertIn("998P160", item.toolTip(6))
        self.assertEqual(window._validation_results["STRUCT6::2C11531G03"].status, "FAIL")
        self.assertEqual(window._validation_results["STRUCT8::2C11531G03"].status, "PASS")


class CodeMappingTests(unittest.TestCase):
    def test_index2_failure_does_not_replace_checklist_with_empty_rows(self):
        window = MagicMock()
        window._find_named_pdf.side_effect = ["bom.pdf", None]
        window._read_pdf.return_value = SimpleNamespace(text="BOM")
        window._pdf_parameter_sources.return_value = []
        window.test_scope_checkbox.isChecked.return_value = False
        window._load_code_mappings.return_value = ({"TEST": object()}, "local checklist")
        window._excel_com.get_index2_map.side_effect = RuntimeError("Excel dialog is open")
        with patch.object(QA2, "parse_structured_bom", return_value=[("1", "Structure", [])]), \
                patch.object(QA2.os, "listdir", return_value=[]), \
                patch.object(QA2.QMessageBox, "warning") as warning:
            QA2.AuditReviewDemo._run_automation(window, "job")
        window.build_bom_checklist.assert_not_called()
        self.assertEqual(warning.call_args.args[1], "Index2 Unavailable")
        self.assertIn("Close any Excel error dialogs", warning.call_args.args[2])

    def test_parameter_module_is_used_without_excel_or_local_workbook(self):
        window = SimpleNamespace(_excel_com=MagicMock(), _set_status=MagicMock())
        with patch.object(QA2, "read_local_sheet_values") as local_read, \
                patch.object(QA2, "load_workbook") as open_workbook:
            mappings, note = QA2.AuditReviewDemo._load_code_mappings(window)
        local_read.assert_not_called()
        open_workbook.assert_not_called()
        window._excel_com.read_sheets_values.assert_not_called()
        self.assertIn("WCUTO22XMAIN", mappings)
        self.assertEqual(window._common_parameters, tuple(QA2.checklist_parameters.COMMON_PARAMETERS))
        self.assertEqual(QA2.EXCEL_PARAMETERS, window._common_parameters)
        self.assertIn("checklist_parameters.py", note)

    def test_updated_parameter_lists_are_loaded_from_the_module(self):
        rows = [("Name", "Code", "Parameter 1"), ("New family", "NEWCODE", "New Parameter")]
        window = SimpleNamespace(_excel_com=MagicMock(), _set_status=MagicMock())
        with patch.object(QA2.checklist_parameters, "CODE_WISE_PARAMETERS", rows), \
                patch.object(QA2.checklist_parameters, "COMMON_PARAMETERS", ["New Parameter"]):
            mappings, _note = QA2.AuditReviewDemo._load_code_mappings(window)
        self.assertEqual(tuple(mappings), ("NEWCODE",))
        self.assertEqual(mappings["NEWCODE"].parameters[0], "New Parameter")
        self.assertEqual(window._common_parameters, ("New Parameter",))
        window._excel_com.read_sheets_values.assert_not_called()

    def test_parameter_definitions_have_no_duplicate_codes_or_names(self):
        rows = QA2.checklist_parameters.CODE_WISE_PARAMETERS
        self.assertEqual(len(QA2.parse_code_mappings(rows)), len(rows) - 1)
        names = QA2.checklist_parameters.COMMON_PARAMETERS
        self.assertEqual(len(names), len(set(names)))

    def test_open_only_checklist_read_never_launches_a_workbook(self):
        session = QA2.ExcelComSession()
        with patch.object(session, "_get_app", return_value=SimpleNamespace(Workbooks=[])), \
                patch.object(session, "_get_index2_workbook") as index2, \
                patch.object(session, "_open_with_sign_in") as launch:
            with self.assertRaisesRegex(RuntimeError, "Open .* in Excel"):
                session.read_sheets_values(QA2.CHECKLIST_LIVE_URLS[0], (QA2.CODE_SHEET_NAME,), open_if_missing=False)
        index2.assert_not_called()
        launch.assert_not_called()

    def test_sheet1_rows_map_by_normalized_code(self):
        rows = (
            ("Name", "Code", "Parameter 1", "Parameter 2"),
            ("Bus Ground", "GND BOND", "Amperage", "Assembly No.", None),
            ("Utilities", "CTWISPUB", "Amperage"),
        )
        mappings = QA2.parse_code_mappings(rows)

        mapping = mappings[QA2.normalize_code("gndbond")]
        self.assertEqual(mapping.name, "Bus Ground")
        self.assertEqual(mapping.code, "GND BOND")
        self.assertEqual(mapping.parameters, ("Amperage", "Assembly No.", "", "", ""))
        self.assertEqual(mappings["CTWISPUB"].parameters, ("Amperage", "", "", "", ""))


class GroupValidationTests(unittest.TestCase):
    values = (
        ("8708C64 ASSEMBLY", None, None),
        (None, None, None),
        ("Amperage", "Material", "Group"),
        (800, "Copper", "G01"),
        (1200, "Copper", "G02"),
        (1200, "Aluminum", "G03"),
        (None, None, None),
        ("Amperage", "Group", None),
        (400, "G99", None),
    )

    def tables(self):
        return QA2.WorkbookParser.extract_tables(self.values, "BUS")

    def test_parser_stops_tables_at_blank_rows(self):
        tables = self.tables()
        self.assertEqual([len(table.rows) for table in tables], [3, 1])

    def test_group_found_from_all_matching_parameters(self):
        group, _ = QA2.GroupDeterminationEngine.identify_group(
            self.tables(), "8708C64", {"Amperage": "1200A", "Material": "aluminum"},
        )
        self.assertEqual(group, "G03")

    def test_ambiguous_parameters_do_not_guess_a_group(self):
        group, message = QA2.GroupDeterminationEngine.identify_group(
            self.tables(), "8708C64", {"Amperage": "1200"},
        )
        self.assertEqual(group, "")
        self.assertIn("Multiple groups", message)

    def test_table_for_other_base_is_ignored(self):
        group, _ = QA2.GroupDeterminationEngine.identify_group(
            self.tables(), "9999C99", {"Amperage": "800"},
        )
        self.assertEqual(group, "")

    def test_compare_reports_pass_fail_and_unresolved(self):
        compare = QA2.ValidationEngine.compare
        self.assertEqual(compare("G03", "g03").status, "PASS")
        self.assertEqual(compare("G03", "G02").status, "FAIL")
        self.assertEqual(compare("G03", "", "why").status, "UNRESOLVED")
        self.assertEqual(compare("", "G02").status, "UNRESOLVED")


class WorkbookValueTests(unittest.TestCase):
    @staticmethod
    def wcctaux3_sheet():
        return (("WCCTAUX3", (
            ("WEST COAST UGPS AUX CROSS BUS APPLICATION A",),
            ("REF - 2C11531 (51W STRUCTURE)",),
            ("AMPS", "COPPER", "COPPER"),
            (None, "HEAT TEST", "DENSITY"),
            ("2500/3000", "G02", "G03"),
        )),)

    def test_section_bus_override_drives_independent_group_comparison(self):
        window = SimpleNamespace(
            _parameter_repository=QA2.ParameterRepository(),
            parameter_variables=QA2.ExcelParameterVariables.from_sources([
                "Bus Amps: 3000\nBus Material: Silver Plated Copper\nHeat Test",
            ]),
            _section_bus_requirements=QA2.section_bus_requirements(
                "SECTION 1,2,3,4,5,7&8 NEEDS TO BE DENSITY RATED BUS (998P160)", "outline.dwg",
            ),
            _assembly_links={"2C11531G03": "WCCTAUX3.xlsx", "2C11531G02": "WCCTAUX3.xlsx"},
            _excel_com=SimpleNamespace(get_sheet_values=lambda link: self.wcctaux3_sheet()),
        )
        window._resolved_value = lambda *args: QA2.AuditReviewDemo._resolved_value(window, *args)
        mapping = QA2.parse_code_mappings((
            ("Name", "Code", "Parameter 1", "Parameter 2", "Parameter 3"),
            ("Bus", "WCTAUX5A", "Amperage", "Material", "Heat Test/Density"),
        ))["WCTAUX5A"]
        validate = QA2.AuditReviewDemo.validate_assembly
        for section, style, expected_actual, status in ((8, "2C11531G03", "G03", "PASS"), (6, "2C11531G03", "G02", "FAIL"), (8, "2C11531G02", "G03", "FAIL")):
            with self.subTest(section=section, style=style):
                result = validate(window, style, mapping, "51", "36", (), "", "XBUS,UGPS,25-3000,CU DEN,51W", section)
                self.assertEqual((result.actual_group, result.status), (expected_actual, status))
                if section == 8:
                    self.assertIn("998P160", result.message)
        window._section_bus_requirements[8].add_parameter("Heat Test/Density", "HEAT TEST", "other drawing")
        result = validate(window, "2C11531G03", mapping, "51", "36", (), "", "XBUS,UGPS,25-3000,CU DEN,51W", 8)
        self.assertEqual((result.actual_group, result.status), ("", "UNRESOLVED"))
        self.assertIn("Conflicting section 8", result.message)
        self.assertIsNone(QA2.assembly_bus_requirements(window._section_bus_requirements, 8, "FLAT ROOF,51W"))

    def test_description_selects_cross_bus_table_before_group_parameters(self):
        sheets = [("BUS TABLES", (
            ("SCMM CROSS BUS",),
            ("ASSY DWG", '36" WIDE', '38" WIDE'),
            ("42C7103", "G84", "G85"),
            ("SCMM CROSS BUS MOUNTING",),
            ("ASSY DWG", '36" WIDE', '38" WIDE'),
            ("42C7103", "G14", "G15"),
        ))]
        group, message = QA2.find_grid_group(
            sheets, "42C7103", {"Width (Wide)": "36"},
            description="ASSY, SCMM CROSS BUS, 36W",
        )
        self.assertEqual(group, "G84")
        self.assertIn("SCMM CROSS BUS", message)

    def test_description_expands_mounting_abbreviation(self):
        matches = [
            QA2.GridMatch("Sheet1", (), heading, (), "", group)
            for heading, group in (("WRP CROSS BUS", "G01"), ("WRP CROSS BUS MOUNTING", "G05"))
        ]
        selected, error = QA2.select_grid_tables(matches, {}, "SWBC, WRP CROSS BUS MTG 36W")
        self.assertEqual([match.group for match in selected], ["G05"])
        self.assertEqual(error, "")

    def test_single_mounting_table_needs_no_description_selector(self):
        matches = [QA2.GridMatch("Sheet1", (), "CROSS BUS MOUNTING - AUXILIARY (TOP)", (), "", "G05")]
        self.assertEqual(QA2.select_grid_tables(matches, {}), (matches, ""))

    def test_wrp_full_assembly_columns_keep_height_and_channel_choices(self):
        sheets = [("WRPCBMTG", (
            (None, "WRP CROSS BUS MOUNTING"),
            (None,),
            (None, '90" HIGH STRUCTURE', None, None, '78" HIGH STRUCTURE'),
            (None, "STRUCTURE", "SPLIT CHANNEL", "FULL CHANNEL *", "STRUCTURE", "SPLIT CHANNEL", "FULL CHANNEL *"),
            (None, "WIDTH", "ASSEMBLY", "ASSEMBLY", "WIDTH", "ASSEMBLY", "ASSEMBLY"),
            (None, "36", "6573C44G05", "6573C44G13", "36", "6573C44G22", "6573C44G30"),
        ))]
        wanted = {"Width (Wide)": "36"}
        description = "SWBC, WRP CROSS BUS MTG 36W"
        group, message = QA2.find_grid_group(sheets, "6573C44", wanted, description=description)
        self.assertEqual(group, "")
        self.assertIn("Multiple groups", message)
        wanted.update({"Height": "90", "Channel Type": "SPLIT CHANNEL"})
        self.assertEqual(QA2.find_grid_group(sheets, "6573C44", wanted, description=description)[0], "G05")

    def test_scmm_cross_bus_inherits_only_its_local_parent_headers(self):
        sheet = (
            ("SINGLE PHASE CROSS BUS ASSEMBLY - BUS POSITION 1 ONLY",),
            ("ASSY DWG", "AMPS", "26W"),
            (None, None, "Density"),
            (None, None, "CUAG"),
            ("42C7102", "2500A", "G38"),
            ("THREE PHASE CROSS BUS ASSEMBLY - BUS POSITION 1 ONLY",),
            ("42C7103", "2500A", "G38"),
            (None,),
            ("SINGLE PHASE CROSS BUS ASSEMBLY - BUS POSITION 1 ONLY",),
            ("ASSY DWG", "AMPS", "36W", None),
            (None, None, "Density", "Heat Tested"),
            (None, None, "CUAG", "CUAG"),
            ("42C7102", "2500A", "G84", "G81"),
            ("THREE PHASE CROSS BUS ASSEMBLY - BUS POSITION 1 ONLY",),
            ("42C7103", "2500A", "G84", "G81"),
        )
        wanted = {"Width (Wide)": "36", "Material": "CUAG", "Heat Test/Density": "DENSITY", "Amperage": "2500"}
        self.assertEqual(QA2.find_grid_group([("SCMM", sheet)], "42C7103", wanted, description="ASSY, SCMM CROSS BUS, 36W")[0], "G84")

    def test_ambiguous_cross_bus_descriptions_do_not_use_group_to_pick_table(self):
        matches = [
            QA2.GridMatch("Sheet1", (), heading, (), "", "G05")
            for heading in ("SCMM CROSS BUS", "WRP CROSS BUS")
        ]
        selected, error = QA2.select_grid_tables(matches, {}, "ASSY, CROSS BUS, 36W")
        self.assertEqual(selected, [])
        self.assertIn("Table selection unresolved", error)

    def test_repeated_group_does_not_combine_incompatible_parameter_cells(self):
        sheets = [("SCMM", (
            ("THREE PHASE CROSS BUS ASSEMBLY",),
            ("ASSY DWG", "AMPS", "36W", None),
            (None, None, "Density", "Heat Tested"),
            (None, None, "CUAG", "CUAG"),
            ("42C7103", "1600A", "G80", "G77"),
            (None, "2500A", "G84", "G80"),
        ))]
        wanted = {"Width (Wide)": "36", "Material": "CUAG", "Heat Test/Density": "DENSITY", "Amperage": "2500"}
        self.assertEqual(QA2.find_grid_group(sheets, "42C7103", wanted, description="ASSY, SCMM CROSS BUS, 36W")[0], "G84")

    def test_numeric_top_cover_axes_do_not_borrow_adjacent_drawing_groups(self):
        sheet = (
            ("TYPE 1 TOP COVER", None, None, None, "BUS DUCT TOP COVER"),
            ("STRUCTURE WIDTH", "ASSEMBLY DRAWING", "DEPTH", None, "STRUCTURE WIDTH", "ASSEMBLY DRAWING", "DEPTH"),
            (None, None, 24, 30, None, None, 30),
            (30, "6575C88", "G10", "G16", 30, "42C9673", "G90"),
            (36, None, "G11", "G17", 36, None, "G91"),
        )
        sheets = [("TOPCOVER", sheet)]
        self.assertEqual(QA2.find_grid_group(sheets, "6575C88", {"Width (Wide)": "30", "Depth": "30"})[0], "G16")
        self.assertEqual(QA2.find_grid_group(sheets, "6575C88", {"Width (Wide)": "36", "Depth": "30"})[0], "G17")
        self.assertEqual(QA2.find_grid_group(sheets, "42C9673", {"Width (Wide)": "30", "Depth": "30"})[0], "G90")
        self.assertEqual(QA2.find_grid_group(sheets, "6575C88", {"Width (Wide)": "45", "Depth": "30"})[0], "")

    def test_mda_structure_dimensions_and_standalone_heat_test(self):
        first = [("6575C01G16", "SWBC, 30 X 30 STRUCTURE ASSY 1 0 F NOLOC", "", "")]
        second = [("6575C01G22", "SWBC, 36 X 30 STRUCTURE ASSY 1 0 F NOLOC", "", "")]
        self.assertEqual(QA2.structure_dimensions_of(first), ("30", "30"))
        self.assertEqual(QA2.structure_dimensions_of(second), ("36", "30"))
        self.assertEqual(QA2.structure_width_of(second), "36")
        variables = QA2.ExcelParameterVariables.from_sources([
            "Bus Amps: 1200\nBus Material: Copper    Heat Test\nEnclosure Type: Type 1",
        ])
        self.assertEqual(variables.Heat_Test_Density, "HEAT TEST")
        self.assertEqual(variables.Material, "COPPER")

    def test_group_inputs_use_saved_common_variables_without_guessing_conflicts(self):
        window = SimpleNamespace(
            _parameter_repository=QA2.ParameterRepository(),
            parameter_variables=QA2.ExcelParameterVariables.from_sources([
                "Bus Amps: 1200\nBus Material: Copper",
            ]),
        )
        resolve = QA2.AuditReviewDemo._resolved_value
        self.assertEqual(resolve(window, "6575C20G06", "Amperage", "", ""), "1200")
        self.assertEqual(resolve(window, "6575C20G06", "Material", "", ""), "COPPER")
        window.parameter_variables = QA2.ExcelParameterVariables({"Amperage": ["1200", "800"]})
        self.assertEqual(resolve(window, "6575C20G06", "Amperage", "", ""), "")
        window.parameter_variables = QA2.ExcelParameterVariables({"Material": ["COPPER"]})
        window._parameter_repository.add_parameter("Material", "Copper", "first")
        window._parameter_repository.add_parameter("Material", "Aluminum", "second")
        self.assertEqual(resolve(window, "6575C20G06", "Material", "", ""), "")

    def test_section_bus_requirements_keep_scope_and_source(self):
        note = "2. HAVE A QUICK CHECK OF BUSES & SPLICES IN SECTION 1,2,3,4,5,7&8 IN DA BOM. IT NEEDS TO BE DENSITY RATED BUS AS PER NEW GUIDELINES (998P160)"
        requirements = QA2.section_bus_requirements(note, "outline.dwg")
        self.assertEqual(set(requirements), {1, 2, 3, 4, 5, 7, 8})
        self.assertEqual(requirements[8].get_parameter("Heat Test/Density"), "DENSITY")
        self.assertIn("998P160", requirements[8].get_entries("Heat Test/Density")[0][1])
        self.assertEqual(QA2.section_bus_requirements("Bus Material: Copper Heat Test", "QuickPrint"), {})
        ranged = QA2.section_bus_requirements("STRUCTURES 2-4 AND 8 MUST USE HEAT TESTED BUS", "drawing")
        self.assertEqual(set(ranged), {2, 3, 4, 8})
        self.assertEqual(ranged[8].get_parameter("Heat Test/Density"), "HEAT TEST")

    def test_saved_bus_position_selects_table_before_width(self):
        sheet = (
            ("CROSS BUS MOUNTING - AUXILIARY (TOP)",),
            ("ASSY DWG", '36" WIDE', '38" WIDE'),
            ("42C2066", "G05", "G06"),
            ("CROSS BUS MOUNTING - AUXILIARY (TOP) AND STANDARD (BOTTOM)",),
            ("ASSY DWG", '36" WIDE', '38" WIDE'),
            ("42C2066", "G14", "G15"),
        )
        sheets = [("AUXILCB", sheet)]
        saved = QA2.ExcelParameterVariables.from_sources(["Bus Position: Auxiliary (Top)"])
        self.assertEqual(saved.Bus_Position, "AUXILIARY (TOP)")
        group, message = QA2.find_grid_group(sheets, "42C2066", {"Width (Wide)": "36"}, saved.values)
        self.assertEqual(group, "G05")
        self.assertIn("AUXILIARY (TOP)", message)
        self.assertEqual(QA2.find_grid_group(sheets, "42C2066", {"Width (Wide)": "45"}, saved.values)[0], "")
        for selectors in ({}, {"Bus Position": ["Bottom"]}, {"Bus Position": ["Auxiliary (Top)", "Bottom"]}):
            group, message = QA2.find_grid_group(sheets, "42C2066", {"Width (Wide)": "36"}, selectors)
            self.assertEqual(group, "")
            self.assertIn("Table selection unresolved", message)
        combined = {"Part Name": ["Auxiliary (Top) and Standard (Bottom)"]}
        self.assertEqual(QA2.find_grid_group(sheets, "42C2066", {"Width (Wide)": "36"}, combined)[0], "G14")
        combined_saved = QA2.ExcelParameterVariables.from_sources([
            "Bus Position: Auxiliary (Top) and Standard (Bottom)",
        ])
        self.assertEqual(QA2.find_grid_group(sheets, "42C2066", {"Width (Wide)": "36"}, combined_saved.values)[0], "G14")
        window = SimpleNamespace(
            parameter_variables=saved,
            _resolved_value=lambda style, name, width, depth: width if name == "Width (Wide)" else "",
            _assembly_links={"42C2066G14": "AUXMAMTG.xlsx"},
            _excel_com=SimpleNamespace(get_sheet_values=lambda link: sheets),
        )
        mapping = QA2.parse_code_mappings((("Name", "Code", "Parameter"), ("", "AUXMAMTG", "Width (Wide)")))["AUXMAMTG"]
        result = QA2.AuditReviewDemo.validate_assembly(window, "42C2066G14", mapping, "36", "", (), "")
        self.assertEqual((result.actual_group, result.status), ("G05", "FAIL"))

    @staticmethod
    def wcuto22xmain_sheet():
        def row(*cells):
            return tuple(cells) + (None,) * (7 - len(cells))

        return (
            row(None, '400A THRU 1000A'),
            row(None, 'WEST COAST UTILITY COMPARTMENT'),
            row(None, 'BUS TO 22X CHASSIS MAIN (36"W)'),
            row(None, '9" BUS CENTERS (2C11905, 2C11907)'),
            row(None, 'THREE PHASE'),
            row(None, 'CURRENT', 'ASSEMBLY', 'ALUM', 'CUAG', 'CUSN'),
            row(None, 'RATING', 'DRAWING', 'DENSITY', 'DENSITY', 'DENSITY'),
            row(None, '400A, K/PD3(400A) FRAME *', '42C5490', 'G01', 'G02', 'G03'),
            row(None, '800A, MD/PD4 FRAME', '42C5495', 'G01', 'G02', 'G03'),
            row(None, '1000A, N/PD5 FRAME', '42C5496', 'G01', 'G02', 'G03'),
            row(None, 'SINGLE PHASE'),
            row(None, 'CURRENT', 'ASSEMBLY', 'ALUM', 'CUAG', 'CUSN'),
            row(None, 'RATING', 'DRAWING', 'DENSITY', 'DENSITY', 'DENSITY'),
            row(None, '800A, MD/PD4 FRAME', '42C5495', 'G04', 'G05', 'G06'),
        )

    def test_grid_layout_gives_material_density_amperage_and_width(self):
        matches = QA2.find_grid_group_matches(self.wcuto22xmain_sheet(), "400A - 1000A (36W)", "42C5495", "G02")
        self.assertEqual(len(matches), 1)
        self.assertEqual(QA2.grid_parameter_values(matches[0]), {
            "Material": "CUAG", "Heat Test/Density": "DENSITY", "3/1 Phase": "3",
            "Amperage": "800", "Width (Wide)": "36",
        })
        single = QA2.find_grid_group_matches(self.wcuto22xmain_sheet(), "400A - 1000A (36W)", "42C5495", "G05")
        self.assertEqual(QA2.grid_parameter_values(single[0])["3/1 Phase"], "1")

    def test_actual_group_comes_from_description_values(self):
        sheets = [("400A - 1000A (36W)", self.wcuto22xmain_sheet())]
        wanted = QA2.description_parameter_values("800A CUAG DEN 3PH 1 0 F NOLOC")
        self.assertEqual(QA2.find_grid_group(sheets, "42C5495", wanted)[0], "G02")
        single = QA2.description_parameter_values("800A CUAG DEN 1PH")
        self.assertEqual(QA2.find_grid_group(sheets, "42C5495", single)[0], "G05")
        group, message = QA2.find_grid_group(sheets, "42C5495", {"Heat Test/Density": "DENSITY"})
        self.assertEqual(group, "")
        self.assertIn("Multiple groups", message)
        self.assertEqual(QA2.description_parameter_values("XBUS,UGPS,25-3000,CU DEN,51W")["Amperage"], "2500-3000")
        self.assertTrue(QA2._grid_value_matches("CUAG", "CU"))
        self.assertTrue(QA2._grid_value_matches("2500/3000", "2500-3000"))
        self.assertFalse(QA2._grid_value_matches("4000", "2500-3000"))

    def test_sheet_is_chosen_by_structure_width(self):
        narrow = QA2.find_grid_group_matches(self.wcuto22xmain_sheet(), "400A - 1000A (36W)", "42C5495", "G02")
        wide = QA2.find_grid_group_matches(self.wcuto22xmain_sheet(), "400A - 1000A (45W)", "42C5495", "G02")
        self.assertEqual(QA2.merge_grid_matches(narrow + wide, "36")["Width (Wide)"], "36")
        self.assertEqual(QA2.merge_grid_matches(narrow + wide, "45")["Width (Wide)"], "45")
        self.assertEqual(QA2.find_grid_group_matches(self.wcuto22xmain_sheet(), "x", "42C9999", "G02"), [])

    def test_merged_headings_and_base_row_above_the_table(self):
        sheet = (
            (None, "1200A THRU 4000A"),
            (None, "1200-2000A"),
            (None, "2C11909", "38 W  x  30D MIN", None, None, None, None, None, "45 W  x  30D MIN"),
            (None, None, "HEAT TEST", None, None, "DENSITY", None, None, "HEAT TEST"),
            (None, "AMP RATING", "CUAG", "AL", "CUSN", "CUAG", "AL", "CUSN", "CUAG"),
            (None, "1200A", "G08", "G02", "G18", "G09", "G03", "G19", "N/A"),
        )
        found = QA2.merge_grid_matches(QA2.find_grid_group_matches(sheet, "Sheet1", "2C11909", "G03"), "38")
        self.assertEqual(found, {
            "Material": "AL", "Heat Test/Density": "DENSITY", "Amperage": "1200",
            "Width (Wide)": "38", "Depth": "30D MIN",
        })
        self.assertEqual(QA2.find_grid_group_matches(sheet, "Sheet1", "2C11999", "G03"), [])

    def test_ugps_row_gives_range_drillings_and_depth(self):
        sheet = (
            (None, "WEST COAST UTILITY"),
            (None, "DRAWING", "AMPACITY", "WIDTH", "BUS", "30D", "36D", "48D"),
            (None, None, None, None, None, "(8 DRILLINGS)", "(10 DRILLINGS)", "(13 DRILLINGS)"),
            (None, "2C11518", "2500-3000A", "51W", "CU DEN", "G02", "G03", "G04"),
        )
        found = QA2.merge_grid_matches(QA2.find_grid_group_matches(sheet, "UGPS", "2C11518", "G03"))
        self.assertEqual(found["Number of Drillings"], "10")
        self.assertEqual(found["Depth"], "36D")
        self.assertEqual(found["Amperage"], "2500-3000")
        self.assertEqual((found["Material"], found["Heat Test/Density"], found["Width (Wide)"]), ("CU", "DENSITY", "51"))

    def test_structure_width_comes_from_the_structure_assembly_row(self):
        items = [
            ("42C5495G02", "800A CUAG DEN 3PH", "", ""),
            ("43C1696G52", 'ASSY.SEISMIC STRUCTURE, 36"W X', "", ""),
        ]
        self.assertEqual(QA2.structure_width_of(items), "36")

    def test_excel_parameters_become_variables_named_after_them(self):
        self.assertEqual(QA2.excel_variable_name("Width (Wide)"), "Width_Wide")
        self.assertEqual(QA2.excel_variable_name("3/1 Phase"), "_3_1_Phase")
        self.assertEqual(
            QA2.excel_variable_name("Service/Non-Service Enterance"), "Service_Non_Service_Enterance",
        )

        text = (
            "480Y/277V 3-PHASE 4-WIRE Bus Amps: 1600 Bus Material: Silver Plated Copper "
            "Heat Test: 65 C Main Breaker - Ind Mtd-3000A 3P (8) EUSERC Press Bolts "
            "TYPE 3R (NONWALK-IN) FLAT ROOF"
        )
        variables = QA2.ExcelParameterVariables.from_sources([text], ["43C1696G50"])

        self.assertEqual(variables.Amperage, "1600")
        self.assertEqual(variables.Material, "SILVER PLATED COPPER")
        self.assertEqual(variables.Voltage, "480Y/277V")
        self.assertEqual(variables._3_1_Phase, "3")
        self.assertEqual(variables.Heat_Test_Density, "65 C")
        self.assertEqual(variables.Number_of_Drillings, "8")
        self.assertEqual(variables.Flat_Roof_Sloped_Roof, "FLAT")
        self.assertEqual(variables.Assembly_No, "43C1696G50")
        self.assertEqual(variables.Height, "")
        self.assertEqual(variables.values["Poles"], ["3"])
        for name in QA2.EXCEL_PARAMETERS:
            self.assertTrue(hasattr(variables, QA2.excel_variable_name(name)), name)
    def test_repeated_header_row_and_xlsx_codes_in_sheet1(self):
        rows = (
            ("Name", "Code", "Parameter 1"),
            (None, "Code", "Parameter 1"),
            (None, "IFSXFMRSEL 225-300kVA.xlsx", "Depth"),
        )
        mappings = QA2.parse_code_mappings(rows)
        self.assertEqual(list(mappings), ["IFSXFMRSEL225300KVA"])
        self.assertEqual(QA2.normalize_code("IFSXFMRSEL 225-300kVA"), "IFSXFMRSEL225300KVA")
    def test_common_parameters_and_code_wise_rows(self):
        common = (("Parameters",), ("Amperage",), ("Part Name",), ("Part Name",), (None,), ("Width (Wide)",))
        self.assertEqual(QA2.parse_common_parameters(common), ("Amperage", "Part Name", "Width (Wide)"))

        rows = (
            ("Name", "Code", "Parameter 1"),
            ("Pull", "CABLE PULL", "Amperage", "Part Name", "A", "B", "C", "Sixth"),
        )
        mapping = QA2.parse_code_mappings(rows)["CABLEPULL"]
        self.assertEqual(mapping.parameters, ("Amperage", "Part Name", "A", "B", "C", "Sixth"))
        self.assertEqual(mapping.aliases, mapping.parameters)

        variables = QA2.ExcelParameterVariables.from_sources(["Bus Amps: 1600"], names=("Amperage", "Part Name"))
        self.assertEqual(tuple(variables.values), ("Amperage", "Part Name"))
        self.assertEqual(variables.Amperage, "1600")

    def test_values_come_from_the_row_of_base_and_group(self):
        values = (
            ("8708C64 ASSEMBLY", None, None),
            ("Ampacity", "Part Name", "Group"),
            (800, "Splice A", "G01"),
            (1200, "Splice B", "G02"),
        )
        tables = QA2.WorkbookParser.extract_tables(values, "SPLICE")
        found = QA2.TableValueExtractor.extract(
            tables, "8708C64", "G02",
            {"Amperage": ("Amperage", "Ampacity"), "Part Name": ("Part Name",), "Depth": ("Depth",)},
        )
        self.assertEqual(found, {"Amperage": "1200", "Part Name": "Splice B"})


class IndexMappingTests(unittest.TestCase):
    def test_local_map_uses_only_column_k_and_accepts_single_base(self):
        from openpyxl import Workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "INDEX2"
        sheet["D3"] = "DECOY1, DECOY2"
        sheet["E3"].hyperlink = "REARCOVER.xlsx"
        sheet["K3"] = "43C1673"
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "Index2.xlsx")
            workbook.save(path)
            self.assertEqual(QA2.build_index2_map(path), {"43C1673": "REARCOVER.xlsx"})
        workbook.close()

    def test_live_map_handles_used_range_offset_and_single_base(self):
        sheet = SimpleNamespace(
            UsedRange=SimpleNamespace(
                Row=10, Column=5, Columns=SimpleNamespace(Count=7),
                Value=(("DECOY1, DECOY2", None, None, None, None, None, "43C1673"),),
            ),
            Hyperlinks=[SimpleNamespace(
                Range=SimpleNamespace(Row=10, Column=5), Address="REARCOVER.xlsx",
            )],
        )
        self.assertEqual(QA2.build_index2_map_from_com(sheet), {"43C1673": "REARCOVER.xlsx"})


class FamilyLayoutTests(unittest.TestCase):
    def lifting_table(self, depth="24D-48D", base="8708C64"):
        return (
            (f"{base} ASSEMBLY", None, None),
            (None, None, None),
            ("Depth", None, "Group"),
            (depth, None, "G02"),
        )

    def test_rear_cover_width_is_not_structure_height(self):
        values = (
            ("Structure Height", "Structure Width", "Assembly", "Group"),
            (90, 9, "43C1673", "G01"),
            (None, 36, "43C1673", "G07"),
        )
        result = QA2.table_dimension_lookup(values, "REARCOVER", "43C1673", "G07")
        self.assertEqual((result.width, result.depth), ("36", ""))

    def test_lifting_assembly_preserves_depth_range(self):
        result = QA2.table_dimension_lookup(self.lifting_table(), "WPFLAT", "8708C64", "G02")
        self.assertEqual((result.width, result.depth), ("", "24D-48D"))

    def test_partial_base_match_is_rejected(self):
        result = QA2.table_dimension_lookup(
            self.lifting_table(base="8708C640"), "WPFLAT", "8708C64", "G02",
        )
        self.assertIsNone(result)

    def test_missing_group_does_not_borrow_dimensions(self):
        result = QA2.table_dimension_lookup(self.lifting_table(), "WPFLAT", "8708C64", "G99")
        self.assertEqual((result.width, result.depth), ("", ""))
        self.assertIn("not supported", result.message)

    def test_bus_properties_are_not_width_and_depth(self):
        values = (
            ("6575C43 - 3PH 3W", None, None, None),
            (None, None, None, None),
            ("OFFSET", "BARS PER PHASE", '3" BUS SIZE', None),
            (None, None, "GROUPS", "AMP RATING"),
            (None, None, None, "HT"),
            (None, None, None, None),
            ('6"', 1, "G02", 800),
        )
        result = QA2.table_dimension_lookup(values, "BUSSPLIC", "6575C43", "G02")
        self.assertEqual((result.width, result.depth), ("", ""))
        self.assertIn("not Width/Depth", result.message)

    def test_conflicting_active_sheets_are_unresolved(self):
        workbook = SimpleNamespace(Name="WPFLAT.xlsx", Worksheets=[
            SimpleNamespace(Name="WPFLAT-Steel", UsedRange=SimpleNamespace(Value=self.lifting_table())),
            SimpleNamespace(Name="WPFLAT-Galvanized Steel", UsedRange=SimpleNamespace(Value=self.lifting_table("30D-48D"))),
        ])
        session = QA2.ExcelComSession()
        with patch.object(session, "get_family_workbook", return_value=workbook):
            result = session.lookup_dimensions("WPFLAT.xlsx", "BASICSTR-NEMA 3R", "8708C64", "G02")
        self.assertEqual((result.width, result.depth), ("", ""))
        self.assertIn("selection unresolved", result.message)

    def test_equal_active_sheets_agree_and_obsolete_sheet_is_ignored(self):
        workbook = SimpleNamespace(Name="WPFLAT.xlsx", Worksheets=[
            SimpleNamespace(Name="WPFLAT-Steel", UsedRange=SimpleNamespace(Value=self.lifting_table())),
            SimpleNamespace(Name="WPFLAT-Galvanized Steel", UsedRange=SimpleNamespace(Value=self.lifting_table())),
            SimpleNamespace(Name="Obsolete__Old", UsedRange=SimpleNamespace(Value=self.lifting_table("30D-48D"))),
        ])
        session = QA2.ExcelComSession()
        with patch.object(session, "get_family_workbook", return_value=workbook):
            result = session.lookup_dimensions("WPFLAT.xlsx", None, "8708C64", "G02")
        self.assertEqual((result.width, result.depth), ("", "24D-48D"))
        self.assertIn("WPFLAT-Steel", result.source)
        self.assertIn("WPFLAT-Galvanized Steel", result.source)
        self.assertNotIn("Obsolete", result.source)


class ChecklistConfigurationTests(unittest.TestCase):
    def test_checklist_summary_reports_found_and_missing_parameters(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "checklist.xlsx")
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "MCOE Checklist"
            for column, label in enumerate(["REARCOVER", "Assembly No.", "Width (Wide)", "Depth", "Material"], 2):
                sheet.cell(2, column, label)
            workbook.save(path)
            workbook.close()

            configuration = QA2.ConfigurationRepository(path)
            parameters = QA2.ParameterRepository()
            parameters.add_parameter("Material", "Steel", "bom.pdf")
            window = SimpleNamespace(
                _assembly_families={"43C1673G07": "REARCOVER"},
                _configuration_repository=configuration,
                _parameter_repository=parameters,
                enclosure_type="3R",
            )
            summary, detail = QA2.AuditReviewDemo.checklist_summary(window, "43C1673G07", "36", "")

            self.assertEqual(summary, "REARCOVER: 3/4 parameters found")
            self.assertIn("Missing: Depth", detail)
            self.assertIn("Material = Steel", detail)

            window._assembly_families = {}
            self.assertEqual(
                QA2.AuditReviewDemo.checklist_summary(window, "6575C43G02", "", "")[0],
                "Family not resolved from Index2.",
            )

    def test_link_file_name_handles_urls_and_relative_links(self):
        self.assertEqual(QA2.link_file_name("https://x.sharepoint.com/a%20b/REARCOVER.xlsx"), "REARCOVER.xlsx")
        self.assertEqual(QA2.link_file_name("sub\\WPFLAT.xlsx"), "WPFLAT.xlsx")

    def test_absolute_index2_link_is_opened_without_prefixing_folder(self):
        url = "https://eaton.sharepoint.com/sites/Team/Spreadsheets/REARCOVER.xlsx"
        opened = []
        workbook = SimpleNamespace(Name="REARCOVER.xlsx")
        app = SimpleNamespace(
            Workbooks=SimpleNamespace(Open=lambda path, **_kwargs: opened.append(path) or workbook),
        )
        index2 = SimpleNamespace(Path="https://eaton.sharepoint.com/sites/Team/Spreadsheets")
        session = QA2.ExcelComSession()
        with patch.object(session, "_get_app", return_value=app), \
                patch.object(session, "_find_open_workbook", return_value=None), \
                patch.object(session, "_get_index2_workbook", return_value=index2):
            self.assertIs(session.get_family_workbook(url), workbook)
        self.assertEqual(opened, [url])

    def create_checklist(self, path):
        from openpyxl import Workbook

        workbook = Workbook()
        checklist = workbook.active
        checklist.title = "MCOE Checklist"
        database = workbook.create_sheet("Database_Index2")
        checklist["A2"] = "Test Category"
        checklist["B2"] = "TESTCODE"
        checklist["C2"] = "Amperage"
        checklist["D2"] = "Material"
        database["B2"] = "TESTCODE"
        database["C2"] = "Ampacity"
        database["D2"] = "Metal"
        workbook.save(path)
        workbook.close()

    def test_configuration_loads_parameter_aliases(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "checklist.xlsx")
            self.create_checklist(path)
            repository = QA2.ConfigurationRepository(path)

            record = repository.get_configuration("testcode")
            self.assertEqual(record.category, "Test Category")
            self.assertEqual(record.parameters[0].parameter_name, "Amperage")
            self.assertEqual(record.parameters[0].keywords, ("Amperage", "Ampacity"))

    def test_parameter_sources_preserve_and_flag_conflicting_values(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "checklist.xlsx")
            self.create_checklist(path)
            configuration = QA2.ConfigurationRepository(path)
            repository = QA2.ParameterRepository()
            repository.merge_sources(
                configuration.extract_parameters("Amperage: 800A", "bom.pdf"),
                configuration.extract_parameters("Ampacity: 800A", "drawing.pdf"),
            )

            self.assertEqual(repository.get_parameter("Amperage"), "800A")
            self.assertEqual(repository.conflict_count, 0)

            repository.merge_sources(
                configuration.extract_parameters("Ampacity: 1000A", "quickprint.pdf"),
            )
            self.assertIsNone(repository.get_parameter("Amperage"))
            self.assertEqual(repository.conflict_count, 1)
            self.assertEqual(repository.value_count, 3)

    def test_checklist_summary_includes_document_evidence(self):
        from openpyxl import Workbook

        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "checklist.xlsx")
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "MCOE Checklist"
            for column, label in enumerate(["REARCOVER", "Material"], 2):
                sheet.cell(2, column, label)
            workbook.save(path)
            workbook.close()

            configuration = QA2.ConfigurationRepository(path)
            parameters = QA2.ParameterRepository()
            parameters.add_parameter("Material", "Copper", "drawing.pdf (page 2, OCR)")
            parameters.add_parameter("Material", "Steel", "assembly.dwg")
            window = SimpleNamespace(
                _assembly_families={"43C1673G07": "REARCOVER"},
                _configuration_repository=configuration,
                _parameter_repository=parameters,
                enclosure_type="3R",
            )

            summary, detail = QA2.AuditReviewDemo.checklist_summary(
                window, "43C1673G07", "", "",
            )

            self.assertEqual(summary, "REARCOVER: 0/1 parameters found")
            self.assertIn("Copper [drawing.pdf (page 2, OCR)]", detail)
            self.assertIn("Steel [assembly.dwg]", detail)

    def test_native_dwg_reader_extracts_text_and_block_attributes(self):
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "checklist.xlsx")
            self.create_checklist(path)
            configuration = QA2.ConfigurationRepository(path)
            document = SimpleNamespace(
                ModelSpace=[SimpleNamespace(ObjectName="AcDbText", TextString="Amperage: 800A")],
                PaperSpace=[SimpleNamespace(ObjectName="AcDbBlockReference", GetAttributes=lambda: [
                    SimpleNamespace(TagString="Material", TextString="Copper"),
                ])],
                Close=lambda save: None,
            )
            documents = MagicMock()
            documents.__iter__.return_value = iter(())
            documents.Open.return_value = document
            app = SimpleNamespace(Documents=documents)
            com_client = SimpleNamespace(
                GetActiveObject=lambda _prog_id: (_ for _ in ()).throw(RuntimeError("not running")),
                Dispatch=lambda _prog_id: app,
            )

            with patch.object(QA2, "win32", com_client):
                result = QA2.DWGFile(os.path.join(folder, "sample.dwg")).extract_parameters(configuration)

            self.assertEqual(result.get_parameter("Amperage"), "800A")
            self.assertEqual(result.get_parameter("Material"), "Copper")

    def test_native_dwg_reader_preserves_existing_document(self):
        path = os.path.abspath("sample.dwg")
        document = SimpleNamespace(
            FullName=path,
            ModelSpace=[SimpleNamespace(ObjectName="AcDbText", TextString="SECTION 8: DENSITY RATED BUS")],
            PaperSpace=[], Close=MagicMock(),
        )
        documents = MagicMock()
        documents.__iter__.return_value = iter((document,))
        client = SimpleNamespace(GetActiveObject=lambda name: SimpleNamespace(Documents=documents))
        with patch.object(QA2, "win32", client):
            self.assertIn("DENSITY RATED BUS", QA2.DWGFile(path).read_text())
        document.Close.assert_not_called()
        documents.Open.assert_not_called()


class PDFExtractionTests(unittest.TestCase):
    def test_scanned_page_uses_ocr_and_retains_page_provenance(self):
        pdf = MagicMock()
        pdf.__enter__.return_value.pages = [SimpleNamespace(extract_text=lambda: "")]
        rendered_pdf = MagicMock()
        rendered_pdf.load_page.return_value.get_pixmap.return_value.tobytes.return_value = b"png"
        image = MagicMock()

        with patch.object(QA2, "pdfplumber", SimpleNamespace(open=lambda _path: pdf)), \
                patch.object(QA2, "fitz", SimpleNamespace(open=lambda _path: rendered_pdf)), \
                patch.object(QA2, "pytesseract", SimpleNamespace(
                    image_to_string=lambda _image: "Material: Copper",
                )), \
                patch.object(QA2, "Image", SimpleNamespace(
                    open=lambda _image_data: image,
                )):
            extraction = QA2.extract_pdf_pages("drawing.pdf")

        self.assertEqual(extraction.text, "Material: Copper")
        self.assertEqual(extraction.ocr_pages, (1,))
        self.assertEqual(extraction.issues, ())
        self.assertEqual(
            QA2.AuditReviewDemo._pdf_parameter_sources(extraction, "drawing.pdf"),
            [("Material: Copper", "drawing.pdf (page 1, OCR)")],
        )
        rendered_pdf.close.assert_called_once_with()

    def test_scanned_page_reports_missing_ocr_dependencies(self):
        pdf = MagicMock()
        pdf.__enter__.return_value.pages = [SimpleNamespace(extract_text=lambda: "")]
        with patch.object(QA2, "pdfplumber", SimpleNamespace(open=lambda _path: pdf)), \
                patch.object(QA2, "fitz", None), \
                patch.object(QA2, "pytesseract", None), \
                patch.object(QA2, "Image", None):
            extraction = QA2.extract_pdf_pages("scan.pdf")

        self.assertEqual(extraction.pages, ("",))
        self.assertEqual(extraction.ocr_pages, ())
        self.assertIn("OCR unavailable", extraction.issues[0])
        self.assertIn("PyMuPDF, pytesseract, Pillow", extraction.issues[0])
        self.assertIn("Tesseract OCR executable", extraction.issues[0])


if __name__ == "__main__":
    unittest.main()