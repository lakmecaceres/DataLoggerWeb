from flask import Flask, render_template, request, jsonify, send_file, send_from_directory, make_response
import os
import io
import re
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import Workbook, load_workbook
import dateutil.parser

# Optional GCS support
GCS_BUCKET = os.getenv("GCS_BUCKET", "").strip()
GCS_ENABLED = bool(GCS_BUCKET)
if GCS_ENABLED:
    from google.cloud import storage

app = Flask(__name__)
app.config['SECRET_KEY'] = 'marmoset'


class DataLogger:
    def __init__(self):
        # Local storage (fallback when GCS not enabled)
        self.config_dir = os.path.join(os.path.expanduser('~'), 'DataLogApp')
        os.makedirs(self.config_dir, exist_ok=True)
        self.counter_file = os.path.join(self.config_dir, 'sample_name_counter.json')

        # GCS client
        self.storage_client = storage.Client() if GCS_ENABLED else None

        self.name_to_code = {
            "Petra": "CJ23.56.001",
            "Croissant": "CJ23.56.002",
            "Nutmeg": "CJ23.56.003",
            "Tank": "CJ23.56.004",
            "JellyBean": "CJ24.56.001",
            "Pringle": "CJ24.56.002",
            "Paarl": "CJ24.56.003",
            "Rambo": "CJ24.56.004",
            "Clack": "CJ24.56.005",
            "Porthos": "CJ24.56.006",
            "Deegan": "CJ24.56.007",
            "Dangerboy": "CJ24.56.008",
            "Hildegard": "CJ24.56.009",
            "Villopoto": "CJ24.56.010",
            "Pathy": "CJ24.56.011",
            "Toki": "CJ24.56.012",
            "Georgia": "CJ24.56.013",
            "Carmichael": "CJ24.56.014",
            "Morel": "CJ24.56.015",
            "Orion": "CJ24.56.016",
            "EllieMae": "CJ24.56.017",
            "Lambert": "CJ24.56.018",
            "Ocean": "CJ25.56.001",
            "Stella": "CJ25.56.002",
            "Wyatt": "CJ25.56.003",
            "Piglet": "CJ25.56.004",
            "Moira": "CJ25.56.005",
            "Willow": "CJ25.56.006",
            "Wren": "CJ25.56.007",
            "Valentino": "CJ25.56.008",
            "Misty": "CJ25.56.009",
            "Link": "CJ25.56.010",
            "Owlette": "CJ25.56.011",
            "Chickpea": "CJ25.56.012",
            "Benedict": "CJ25.56.013",
            "Vera": "CJ25.56.014",
            "Tango": "CJ25.56.015",
            "Paris": "CJ25.56.016",
            "Lapras": "CJ25.56.017"
        }

        self.black_fill = PatternFill(start_color='000000', fill_type='solid')

    # ----------------- User key and object names -----------------

    def _safe_user_key(self, name: str) -> str:
        name = (name or "").strip()
        if not name:
            return "unknown"
        return re.sub(r'[^A-Za-z0-9_-]+', '_', name) or "unknown"

    # ----------------- Workbook storage helpers (GCS/local) -----------------

    def _new_object_name(self, user_key: str) -> str:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"logs/{user_key}/{user_key}_krienen_data_log_{ts}.xlsx"

    def _load_pointer(self, user_key: str):
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(f"pointers/{user_key}.json")
            if blob.exists():
                try:
                    data = json.loads(blob.download_as_text())
                    if isinstance(data, dict) and data.get("object"):
                        return data["object"]
                except Exception:
                    pass
        # local fallback mapping (optional)
        mapping = self._load_local_meta().get("current_log_objects", {})
        return mapping.get(user_key)

    def _save_pointer(self, user_key: str, object_name: str):
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(f"pointers/{user_key}.json")
            blob.upload_from_string(json.dumps({"object": object_name}, indent=2), content_type="application/json")
        meta = self._load_local_meta()
        meta.setdefault("current_log_objects", {})[user_key] = object_name
        self._save_local_meta(meta)

    def _download_workbook(self, object_name):
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(object_name)
            if blob.exists():
                data = blob.download_as_bytes()
                wb = load_workbook(io.BytesIO(data))
                return wb, blob.generation
            else:
                wb = self._initialize_excel()
                return wb, None
        else:
            local_path = os.path.join(self.config_dir, object_name.replace("/", os.sep))
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            if os.path.exists(local_path):
                wb = load_workbook(local_path)
            else:
                wb = self._initialize_excel()
            return wb, None

    def _upload_workbook(self, wb, object_name, if_generation_match=None):
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(object_name)
            out = io.BytesIO()
            wb.save(out)
            out.seek(0)
            if if_generation_match is not None:
                blob.upload_from_file(
                    out,
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    if_generation_match=if_generation_match
                )
            else:
                blob.upload_from_file(
                    out,
                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            local_path = os.path.join(self.config_dir, object_name.replace("/", os.sep))
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            wb.save(local_path)

    def _download_workbook_bytes(self, object_name):
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(object_name)
            if blob.exists():
                return blob.download_as_bytes()
            else:
                wb = self._initialize_excel()
                out = io.BytesIO()
                wb.save(out)
                return out.getvalue()
        else:
            local_path = os.path.join(self.config_dir, object_name.replace("/", os.sep))
            if os.path.exists(local_path):
                with open(local_path, 'rb') as f:
                    return f.read()
            else:
                wb = self._initialize_excel()
                out = io.BytesIO()
                wb.save(out)
                return out.getvalue()

    # ----------------- Per-user state (local meta fallback only) -----------------

    def _load_local_meta(self):
        if os.path.exists(self.counter_file):
            try:
                with open(self.counter_file, 'r') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_local_meta(self, meta: dict):
        with open(self.counter_file, 'w') as f:
            json.dump(meta, f, indent=4)

    # ----------------- Core utilities -----------------

    def _headers(self):
        return ['krienen_lab_identifier', 'seq_portal', 'elab_link', 'experiment_start_date',
                'mit_name', 'donor_name', 'tissue_name', 'tissue_name_old',
                'dissociated_cell_sample_name', 'facs_population_plan', 'cell_prep_type',
                'study', 'enriched_cell_sample_container_name', 'expc_cell_capture',
                'port_well', 'enriched_cell_sample_name', 'enriched_cell_sample_quantity_count',
                'barcoded_cell_sample_name', 'library_method', 'cDNA_amplification_method',
                'cDNA_amplification_date', 'amplified_cdna_name', 'cDNA_pcr_cycles',
                'rna_amplification_pass_fail', 'percent_cdna_longer_than_400bp',
                'cdna_amplified_quantity_ng', 'cDNA_library_input_ng', 'library_creation_date',
                'library_prep_set', 'library_name', 'tapestation_avg_size_bp',
                'library_num_cycles', 'lib_quantification_ng', 'library_prep_pass_fail',
                'r1_index', 'r2_index', 'ATAC_index']

    def _initialize_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "HMBA"
        ws.append(self._headers())
        for col_num in range(1, len(self._headers()) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(name="Arial", size=10, bold=True)
            cell.alignment = Alignment(horizontal='left')
        return wb

    def convert_date(self, exp_date):
        clean = "".join(c for c in exp_date if c.isdigit())
        if len(clean) == 6:
            try:
                datetime.strptime(clean, '%y%m%d')
                return clean
            except ValueError:
                pass
        try:
            parsed = dateutil.parser.parse(exp_date)
            return parsed.strftime('%y%m%d')
        except ValueError:
            return None

    def convert_index(self, index):
        index = index.strip().upper()
        if len(index) == 3:
            if index[0].isdigit() and index[1].isdigit() and index[2].isalpha():
                return f"{index[2]}{index[0]}{index[1]}"
            elif index[0].isalpha() and index[1].isdigit() and index[2].isdigit():
                return index
        elif len(index) == 2:
            if index[0].isdigit() and index[1].isalpha():
                return f"{index[1]}0{index[0]}"
            elif index[0].isalpha() and index[1].isdigit():
                return f"{index[0]}0{index[1]}"
        return None

    def pad_index(self, index):
        if len(index) == 2 and index[0].isalpha() and index[1].isdigit():
            return f"{index[0]}0{index[1]}"
        return index

    # ----------------- Sheet-derived state helpers -----------------

    def _sheet_date_chip_usage(self, ws, current_date):
        """
        Scan the sheet for rows with experiment_start_date == current_date
        and build a map of chip -> highest used well for that date.
        barcoded_cell_sample_name format: 'P####_##'
        """
        headers = [cell.value for cell in ws[1]]
        date_col = headers.index('experiment_start_date') + 1
        bcsn_col = headers.index('barcoded_cell_sample_name') + 1

        chips_map = {}  # chip_str -> used_wells (int)
        last_chip = None
        last_used = 0

        for row_idx in range(2, ws.max_row + 1):
            date_val = ws.cell(row=row_idx, column=date_col).value
            if date_val != current_date:
                continue
            name_val = ws.cell(row=row_idx, column=bcsn_col).value
            if not name_val or not isinstance(name_val, str):
                continue
            m = re.match(r'^P(\d{4})_(\d)$', name_val)
            if not m:
                continue
            chip = int(m.group(1))
            well = int(m.group(2))
            chips_map[str(chip)] = max(well, int(chips_map.get(str(chip), 0)))
        if chips_map:
            # Highest chip that has entries on this date
            last_chip = max(int(c) for c in chips_map.keys())
            last_used = int(chips_map[str(last_chip)])
        return chips_map, last_chip, last_used

    def _next_amp_name(self, ws, amp_prefix, amp_date):
        """
        Determine the next amplified_cdna_name by scanning existing rows in the sheet:
        pattern: f"{amp_prefix}_{amp_date}_{batch}_{letter}"
        letter cycles A..H; after H, batch increments.
        """
        headers = [cell.value for cell in ws[1]]
        amp_name_col = headers.index('amplified_cdna_name') + 1

        last_batch = 0
        last_letter = None  # 'A'..'H'

        for row_idx in range(2, ws.max_row + 1):
            val = ws.cell(row=row_idx, column=amp_name_col).value
            if not val or not isinstance(val, str):
                continue
            # Example: APLCTX_251001_1_G
            m = re.match(rf'^{re.escape(amp_prefix)}_{re.escape(amp_date)}_(\d+)_([A-H])$', val)
            if not m:
                continue
            b = int(m.group(1))
            L = m.group(2)
            if b > last_batch or (b == last_batch and (last_letter is None or L > last_letter)):
                last_batch = b
                last_letter = L

        if last_batch == 0:
            # No prior reactions for this amp date
            return f"{amp_prefix}_{amp_date}_1_A"

        # Compute next letter/batch
        letters = 'ABCDEFGH'
        if last_letter is None:
            # Shouldn't happen but default safely
            return f"{amp_prefix}_{amp_date}_{last_batch}_A"
        idx = letters.index(last_letter)
        if idx < 7:
            next_letter = letters[idx + 1]
            next_batch = last_batch
        else:
            next_letter = 'A'
            next_batch = last_batch + 1

        return f"{amp_prefix}_{amp_date}_{next_batch}_{next_letter}"

    # ----------------- Business logic -----------------

    def process_form_data(self):
        # Import heavy modules only when needed
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        import pyperclip

        # Load or create workbook
        if self.workbook_path and os.path.exists(self.workbook_path):
            workbook = load_workbook(self.workbook_path)
        else:
            workbook = self.initialize_excel()

        worksheet = workbook.active

        # Fix: Properly detect the actual last row with content
        last_row_with_content = 1
        for row_idx in range(1, worksheet.max_row + 1):
            row_has_content = False
            for cell in worksheet[row_idx]:
                if cell.value is not None:
                    row_has_content = True
                    break
            if row_has_content:
                last_row_with_content = row_idx

        current_row = last_row_with_content + 1

        # Get form values
        current_date = self.convert_date(self.date_input.text())
        mit_name_input = self.marmoset_input.currentText()
        mit_name = "cj" + mit_name_input
        donor_name = self.name_to_code[mit_name_input]

        # Process slab and hemisphere
        slab = self.slab_input.text().strip()
        hemisphere = self.hemisphere_input.currentText().split()[0].upper()
        if hemisphere == "RIGHT":
            slab = str(int(slab) + 40).zfill(2)
        elif hemisphere == "BOTH":
            slab = str(int(slab) + 90).zfill(2)
        else:
            slab = slab.zfill(2)

        # Modified to handle alphanumeric tile values
        tile_value = self.tile_input.text().strip()
        if tile_value.isdigit():
            tile = str(int(tile_value)).zfill(2)
        else:
            tile = tile_value

        # Process tile location
        tile_location_abbr = self.tile_location_input.currentText()

        # Sort method and FACS population
        sort_method = self.sort_method_input.currentText()
        sort_method = sort_method.upper() if sort_method.lower() == "dapi" else sort_method

        if sort_method.lower() == "pooled":
            facs_population = self.facs_population_input.text()
        elif sort_method.lower() == "unsorted":
            facs_population = "no_FACS"
        else:
            facs_population = "DAPI"

        # Get reaction number and update counters
        rxn_number = int(self.rxn_number_input.text())

        # ==========================================
        # UPDATED LOGIC FOR COLUMN R (PXXXX counter)
        # ==========================================
        if current_date not in self.counter_data["date_info"]:
            p_number = self.counter_data.get("next_counter", 90)
            self.counter_data["date_info"][current_date] = {
                "p_number": p_number,
                "total_reactions": 0
            }
            # Increment the global counter for the *next* new date
            self.counter_data["next_counter"] = p_number + 1

        date_entry = self.counter_data["date_info"][current_date]
        existing_total = date_entry["total_reactions"]
        p_number = date_entry["p_number"]

        # Calculate port wells (the _X suffix)
        port_wells = []
        for x in range(rxn_number):
            port_well = existing_total + x + 1
            port_wells.append((p_number, port_well))

        # Update counters
        date_entry["total_reactions"] = existing_total + rxn_number
        # ==========================================

        # Process indices
        atac_indices = [self.convert_index(index) for index in self.atac_indices_input.text().split(",")]
        atac_indices = [self.pad_index(index) for index in atac_indices]

        rna_indices = [self.convert_index(index) for index in self.rna_indices_input.text().split(",")]
        rna_indices = [self.pad_index(index) for index in rna_indices]

        # Initialize common values
        seq_portal = "no"
        elab_link = pyperclip.paste()
        tissue_name = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"
        dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Multiome'
        cell_prep_type = "nuclei"

        sorting_status = "PS" if sort_method.lower() in ["pooled", "dapi"] else "PN"
        sorter_initials = self.sorter_initials_input.text().strip().upper()
        enriched_cell_sample_container_name = f"MPXM_{current_date}_{sorting_status}_{sorter_initials}"

        # Get study name
        study = "HMBA_CjAtlas_Subcortex" if self.project_input.currentText() == "HMBA_CjAtlas_Subcortex" else self.project_name_input.text()

        # Process the data for each reaction and modality
        dup_index_counter = {}
        headers = [cell.value for cell in worksheet[1]]

        for x in range(rxn_number):
            p_number, port_well = port_wells[x]
            barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

            for modality in ["RNA", "ATAC"]:
                self.write_modality_data(
                    worksheet, current_row, modality, x,
                    current_date, mit_name, slab, tile, sort_method,
                    port_well, barcoded_cell_sample_name,
                    sorting_status, sorter_initials,
                    tissue_name, dissociated_cell_sample_name,
                    enriched_cell_sample_container_name,
                    study, seq_portal, elab_link,
                    facs_population, cell_prep_type,
                    rna_indices, atac_indices,
                    headers, dup_index_counter,
                    donor_name
                )
                current_row += 1

        # Save workbook and counter data
        workbook.save(self.workbook_path)
        with open(self.COUNTER_FILE, 'w') as f:
            json.dump(self.counter_data, f, indent=4)

    def write_modality_data(self, worksheet, current_row, modality, x, current_date, mit_name, slab, tile, sort_method,
                            port_well, barcoded_cell_sample_name, form_data, tissue_name_base, rna_indices,
                            atac_indices, headers, dup_index_counter, donor_name,
                            project=None, combined_slab_label=None, slab_count=1):

        # Identifier slab/tile formatting
        if project in {"HMBA_CjAtlas_Cortex", "HMBA_Aim4"} and combined_slab_label and slab_count > 1:
            unpadded = []
            for s in combined_slab_label.split('_'):
                try:
                    unpadded.append(str(int(s)))
                except ValueError:
                    unpadded.append(s)
            slab_part = f"Slabs_{'_'.join(unpadded)}"
        else:
            slab_part = f"Slab{int(slab)}"
        tile_part = f"Tile{int(tile)}" if str(tile).isdigit() else tile

        krienen_lab_identifier = (
            f"{current_date}_HMBA_{mit_name}_{slab_part}_{tile_part}_{sort_method}_{modality}{x + 1}"
        )

        experimenter_initials = form_data['sorter_initials'].strip().upper()
        sorting_status = "PS" if sort_method.lower() in ["pooled", "dapi"] else "PN"
        tissue_name = tissue_name_base

        if project == "HMBA_Aim4":
            dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Rseq'
            enriched_prefix = "MPTX"
            rna_suffix = "TX"
        else:
            dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Multiome'
            enriched_prefix = "MPXM"
            rna_suffix = "XR"
        atac_suffix = "XA"

        enriched_cell_sample_container_name = f"{enriched_prefix}_{current_date}_{sorting_status}_{experimenter_initials}"
        enriched_cell_sample_name = f'{enriched_prefix}_{current_date}_{sorting_status}_{experimenter_initials}_{port_well}'

        study = form_data.get('project', '')

        seq_portal = "no"
        elab_link = form_data.get('elab_link', '')
        facs_population = form_data.get('facs_population', 'no_FACS')
        cell_prep_type = "nuclei"

        library_prep_date = (self.convert_date(form_data['rna_prep_date']) if modality == "RNA"
                             else self.convert_date(form_data['atac_prep_date']))

        if modality == "RNA":
            if project == "HMBA_Aim4":
                library_method = "10xV4"
                library_type_suffix = rna_suffix
            else:
                library_method = "10xMultiome-RSeq"
                library_type_suffix = rna_suffix
            library_type = f"LP{experimenter_initials}{library_type_suffix}"
            library_index = rna_indices[x]

            cdna_concentration = float(form_data['cdna_concentration'].split(',')[x])
            cdna_amplified_quantity = cdna_concentration * 40
            cdna_library_input = cdna_amplified_quantity * 0.25
            percent_cdna_400bp = float(form_data['percent_cdna_400bp'].split(',')[x])
            rna_concentration = float(form_data['rna_lib_concentration'].split(',')[x])
            lib_quant = rna_concentration * 35

            cdna_pcr_cycles = int(form_data['cdna_pcr_cycles'].split(',')[x])
            rna_size = int(form_data['rna_sizes'].split(',')[x])
            library_cycles = int(form_data['library_cycles_rna'].split(',')[x])
        else:
            library_method = "10xMultiome-ASeq"
            library_type = f"LP{experimenter_initials}{atac_suffix}"
            library_index = atac_indices[x]

            atac_concentration = float(form_data['atac_lib_concentration'].split(',')[x])
            lib_quant = atac_concentration * 20

            atac_size = int(form_data['atac_sizes'].split(',')[x])
            library_cycles = int(form_data['library_cycles_atac'].split(',')[x])

            cdna_concentration = None
            cdna_amplified_quantity = None
            cdna_library_input = None
            percent_cdna_400bp = None
            cdna_pcr_cycles = None
            rna_size = None

        key = (library_type, library_prep_date, library_index)
        dup_index_counter[key] = dup_index_counter.get(key, 0) + 1
        library_prep_set = f"{library_type}_{library_prep_date}_{dup_index_counter[key]}"
        library_name = f"{library_prep_set}_{library_index}"

        expected_cell_capture = int(form_data['expected_recovery'])
        concentration = float(form_data['nuclei_concentration'].replace(",", ""))
        volume = float(form_data['nuclei_volume'])
        enriched_cell_sample_quantity_count = round(concentration * volume)

        row_data = [
            krienen_lab_identifier,
            seq_portal,
            elab_link,
            current_date,
            mit_name,
            donor_name,
            tissue_name,
            None,
            dissociated_cell_sample_name,
            facs_population,
            cell_prep_type,
            study,
            enriched_cell_sample_container_name,
            expected_cell_capture,
            port_well,
            enriched_cell_sample_name,
            enriched_cell_sample_quantity_count,
            barcoded_cell_sample_name,
            library_method,
            ("10xV4" if (modality == "RNA" and project == "HMBA_Aim4")
             else "10xMultiome-RSeq" if modality == "RNA" else None),
            self.convert_date(form_data['cdna_amp_date']) if modality == "RNA" else None,
            None,
            cdna_pcr_cycles if modality == "RNA" else None,
            "Pass" if modality == "RNA" else None,
            percent_cdna_400bp if modality == "RNA" else None,
            cdna_amplified_quantity if modality == "RNA" else None,
            cdna_library_input if modality == "RNA" else None,
            library_prep_date,
            library_prep_set,
            library_name,
            rna_size if modality == "RNA" else atac_size,
            library_cycles,
            lib_quant,
            "Pass",
            f"SI-TT-{rna_indices[x]}_i7" if modality == "RNA" else None,
            f"SI-TT-{rna_indices[x]}_b(i5)" if modality == "RNA" else None,
            f"SI-NA-{atac_indices[x]}" if modality == "ATAC" else None
        ]

        # amplified_cdna_name: derive next value by scanning the sheet, not local state
        # ... [Keep the top of write_modality_data the same] ...

        # Handle amplified_cdna_name for RNA with fixed logic for batch counting
        if modality == "RNA":
            cdna_amp_date = self.convert_date(self.cdna_amp_date_input.text())
            if not cdna_amp_date:
                cdna_amp_date = current_date  # Fallback if empty

            # Create a unique key strictly tied to the date
            amp_date_key = f"amp_{cdna_amp_date}"

            # Initialize counter for new dates
            if amp_date_key not in self.counter_data["amp_counter"]:
                self.counter_data["amp_counter"][amp_date_key] = 0

            # Get current counter for this date
            reaction_count = self.counter_data["amp_counter"][amp_date_key]

            # Math for A-H and batch incrementing
            letter = chr(65 + (reaction_count % 8))  # A through H
            batch_num_for_amp = (reaction_count // 8) + 1  # 1, then 2, then 3...

            row_data[21] = f"APLCXR_{cdna_amp_date}_{batch_num_for_amp}_{letter}"

            # Increment the counter for the next sample on this date
            self.counter_data["amp_counter"][amp_date_key] += 1

        # Write to Excel
        for col_num, value in enumerate(row_data, start=1):
            cell = worksheet.cell(row=current_row, column=col_num, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal='left')

            if (modality == "ATAC" and value is None) or (
                    modality == "RNA" and col_num == headers.index('ATAC_index') + 1):
                cell.fill = self.black_fill

        tissue_old_col = headers.index('tissue_name_old') + 1
        worksheet.cell(row=current_row, column=tissue_old_col).fill = self.black_fill


# favicon route (optional, for direct /favicon.ico requests)
@app.route('/favicon.ico')
def favicon():
    resp = make_response(
        send_from_directory(
            os.path.join(app.root_path, 'static'),
            'favicon.ico',
            mimetype='image/vnd.microsoft.icon'
        )
    )
    resp.headers['Cache-Control'] = 'public, max-age=31536000'
    return resp


data_logger = DataLogger()


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/update_counter', methods=['POST'])
def update_counter():
    try:
        request_data = request.json or {}
        # Accept user from JSON or query param
        user_first_name = (request_data.get('user_first_name') or request.args.get('user') or "").strip()
        if not user_first_name:
            return jsonify({'success': False, 'error': 'Missing user_first_name'}), 400

        # Accept counter as number or numeric string
        new_counter_raw = request_data.get('new_counter')
        try:
            new_counter = int(new_counter_raw)
        except (TypeError, ValueError):
            return jsonify({'success': False, 'error': 'Invalid counter value'}), 400
        if new_counter < 0:
            return jsonify({'success': False, 'error': 'Invalid counter value'}), 400

        user_key = data_logger._safe_user_key(user_first_name)

        # If you’re persisting per-user state to GCS:
        # state = data_logger._load_user_state(user_key)
        # state['next_counter'] = new_counter
        # data_logger._save_user_state(user_key, state)

        # If you’re still using the local meta fallback:
        meta = data_logger._load_local_meta()
        states = meta.setdefault('user_states', {})
        state = states.setdefault(user_key, {"next_counter": None, "date_info": {}, "amp_counter": {}})
        state['next_counter'] = new_counter
        data_logger._save_local_meta(meta)

        return jsonify({'success': True, 'new_counter': new_counter})
    except Exception as e:
        # Surface server-side error to help diagnose
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/submit', methods=['POST'])
def submit_data():
    try:
        form_data = request.json
        required_fields = ['user_first_name', 'date', 'marmoset', 'slab', 'tile', 'hemisphere', 'tile_location', 'sort_method',
                           'rxn_number', 'sorter_initials']
        for field in required_fields:
            if not form_data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'})
        success = data_logger.process_form_data(form_data)
        return jsonify({'success': True, 'message': 'Data saved successfully!'}) if success else jsonify({'success': False, 'error': 'Failed to process data'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/download')
def download_excel():
    try:
        user_first_name = (request.args.get('user') or "").strip()
        if not user_first_name:
            return jsonify({'error': 'Missing user name in query parameter ?user='}), 400
        user_key = data_logger._safe_user_key(user_first_name)
        object_name = data_logger._load_pointer(user_key)
        if not object_name:
            object_name = data_logger._new_object_name(user_key)
            data_logger._save_pointer(user_key, object_name)
        data = data_logger._download_workbook_bytes(object_name)
        filename = os.path.basename(object_name)
        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500