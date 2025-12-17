from flask import Flask, render_template, request, jsonify, send_file
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
        self.load_counter_data()

    # ----------------- Helpers: user key and object names -----------------

    def _safe_user_key(self, name: str) -> str:
        # allow letters, numbers, underscore, dash; fallback to "unknown" if empty
        name = (name or "").strip()
        if not name:
            return "unknown"
        key = re.sub(r'[^A-Za-z0-9_-]+', '_', name)
        return key or "unknown"

    def _new_object_name(self, user_key: str) -> str:
        ts = datetime.now().strftime('%YMMDD_%H%M%S')  # keep consistent sortable naming
        # File per user, include user name in file name
        return f"logs/{user_key}/{user_key}_krienen_data_log_{ts}.xlsx"

    # ----------------- Workbook storage helpers (GCS/local) -----------------

    def _load_pointer(self, user_key: str):
        # Try GCS pointer first
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(f"pointers/{user_key}.json")
            if blob.exists():
                try:
                    data = json.loads(blob.download_as_text())
                    if isinstance(data, dict) and "object" in data and data["object"]:
                        return data["object"]
                except Exception:
                    pass
        # Fallback to local pointer
        mapping = self.counter_data.get("current_log_objects", {})
        return mapping.get(user_key)

    def _save_pointer(self, user_key: str, object_name: str):
        # Save both to GCS and local fallback
        if GCS_ENABLED:
            bucket = self.storage_client.bucket(GCS_BUCKET)
            blob = bucket.blob(f"pointers/{user_key}.json")
            blob.upload_from_string(json.dumps({"object": object_name}, indent=2), content_type="application/json")
        if "current_log_objects" not in self.counter_data:
            self.counter_data["current_log_objects"] = {}
        self.counter_data["current_log_objects"][user_key] = object_name
        with open(self.counter_file, 'w') as f:
            json.dump(self.counter_data, f, indent=4)

    def _download_workbook(self, object_name):
        # Returns (workbook, generation_or_None)
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

    # ----------------- Core utilities -----------------

    def load_counter_data(self):
        if os.path.exists(self.counter_file):
            with open(self.counter_file, 'r') as f:
                try:
                    self.counter_data = json.load(f)
                except json.JSONDecodeError:
                    self.counter_data = {}
        else:
            self.counter_data = {}

        self.counter_data.setdefault("next_counter", None)
        # Per-date chip usage map to support continuing chip only for same date
        self.counter_data.setdefault("date_info", {})
        # Amplification date counters (batch/letter naming)
        self.counter_data.setdefault("amp_counter", {})
        # Pointer mapping per user for local fallback
        self.counter_data.setdefault("current_log_objects", {})

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

    # ----------------- Business logic -----------------

    def process_form_data(self, form_data):
        # Resolve user name (first name) and user key
        user_first_name = (form_data.get('user_first_name') or "").strip()
        user_key = self._safe_user_key(user_first_name)

        # Ensure we have a current per-user log object (file named by creation time and user)
        object_name = self._load_pointer(user_key)
        if not object_name:
            object_name = self._new_object_name(user_key)
            self._save_pointer(user_key, object_name)

        # Load the workbook (with GCS generation for concurrency)
        wb, generation = self._download_workbook(object_name)
        ws = wb.active  # Single "HMBA" sheet

        # Find last row with content
        last_row_with_content = 1
        for row_idx in range(1, ws.max_row + 1):
            if any(cell.value is not None for cell in ws[row_idx]):
                last_row_with_content = row_idx
        current_row = last_row_with_content + 1

        # Parse form inputs
        current_date = self.convert_date(form_data['date'])
        mit_name_input = form_data['marmoset']
        mit_name = "cj" + mit_name_input
        donor_name = self.name_to_code[mit_name_input]
        project = form_data.get('project', '')

        raw_slab = form_data['slab'].strip()
        hemisphere = form_data['hemisphere'].split()[0].upper()

        # Multi-slab for Cortex & Aim4
        if project in {"HMBA_CjAtlas_Cortex", "HMBA_Aim4"}:
            slab_list = [s.strip().zfill(2) for s in raw_slab.split(',') if s.strip()]
            if not slab_list:
                raise ValueError(f"No valid slab numbers provided for {project}")
            combined_slab_label = "_".join(slab_list)
            slab_count = len(slab_list)
            slab = slab_list[0]  # first slab (padded) for single-value fields
        else:
            combined_slab_label = None
            slab_count = 1
            slab = raw_slab
            if hemisphere == "RIGHT":
                slab = str(int(slab) + 40).zfill(2)
            elif hemisphere == "BOTH":
                slab = str(int(slab) + 90).zfill(2)
            else:
                slab = slab.zfill(2)

        # Tile padded elsewhere; first-column identifier will handle text tiles specially
        tile_value = form_data['tile'].strip()
        tile = str(int(tile_value)).zfill(2) if tile_value.isdigit() else tile_value

        tile_location_abbr = form_data['tile_location']
        sort_method = form_data['sort_method']
        sort_method = sort_method.upper() if sort_method.lower() == "dapi" else sort_method

        # FACS population (for the column)
        if sort_method.lower() == "pooled":
            facs_population = form_data['facs_population']
        elif sort_method.lower() == "unsorted":
            facs_population = "no_FACS"
        else:
            facs_population = "DAPI"

        rxn_number = int(form_data['rxn_number'])

        # Chip usage per experiment date (continue only on same date)
        if current_date not in self.counter_data["date_info"]:
            self.counter_data["date_info"][current_date] = {"chips": {}}
        chips_map = self.counter_data["date_info"][current_date]["chips"]  # str chip -> used_wells

        if self.counter_data["next_counter"] is None:
            self.counter_data["next_counter"] = 90

        start_chip = int(self.counter_data["next_counter"])
        chip = start_chip
        used = int(chips_map.get(str(chip), 0))

        # Allocate wells with rollover per chip (max 8)
        assignments = []
        updates = {}
        for _ in range(rxn_number):
            if used == 8:
                updates[str(chip)] = used
                chip += 1
                used = int(chips_map.get(str(chip), 0))
            used += 1
            assignments.append((chip, used))
            updates[str(chip)] = used

        chips_map.update(updates)

        # After submission, keep chip if not full; otherwise advance
        last_chip, last_used = assignments[-1]
        if last_used == 8:
            self.counter_data["next_counter"] = last_chip + 1
        else:
            self.counter_data["next_counter"] = last_chip

        # Indices
        atac_indices = [self.convert_index(i) for i in form_data['atac_indices'].split(",")] if form_data.get('atac_indices') else []
        atac_indices = [self.pad_index(i) for i in atac_indices]
        rna_indices = [self.convert_index(i) for i in form_data['rna_indices'].split(",")] if form_data.get('rna_indices') else []
        rna_indices = [self.pad_index(i) for i in rna_indices]

        dup_index_counter = {}
        headers = [cell.value for cell in ws[1]]

        # Modalities
        modalities = ["RNA"] if project == "HMBA_Aim4" else ["RNA", "ATAC"]

        # Tissue name base (padded slab/tile)
        if project in {"HMBA_CjAtlas_Cortex", "HMBA_Aim4"} and combined_slab_label:
            slab_for_tissue = combined_slab_label
        else:
            slab_for_tissue = slab
        tissue_name_base = f"{donor_name}.{tile_location_abbr}.{slab_for_tissue}.{tile}"

        for x in range(rxn_number):
            p_number, port_well = assignments[x]
            barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

            for modality in modalities:
                self.write_modality_data(
                    ws, current_row, modality, x,
                    current_date, mit_name, slab, tile, sort_method,
                    port_well, barcoded_cell_sample_name,
                    form_data,
                    tissue_name_base=tissue_name_base,
                    rna_indices=rna_indices, atac_indices=atac_indices,
                    headers=headers, dup_index_counter=dup_index_counter,
                    donor_name=donor_name,
                    project=project,
                    combined_slab_label=combined_slab_label,
                    slab_count=slab_count
                )
                current_row += 1

        # Save workbook (with concurrency protection if on GCS) and counters locally
        for attempt in range(3):
            try:
                self._upload_workbook(wb, object_name, if_generation_match=generation if GCS_ENABLED else None)
                break
            except Exception as e:
                if GCS_ENABLED and "precondition" in str(e).lower():
                    wb, generation = self._download_workbook(object_name)
                else:
                    raise

        with open(self.counter_file, 'w') as f:
            json.dump(self.counter_data, f, indent=4)

        return True

    def write_modality_data(self, worksheet, current_row, modality, x, current_date, mit_name, slab, tile, sort_method,
                            port_well, barcoded_cell_sample_name, form_data, tissue_name_base, rna_indices,
                            atac_indices, headers, dup_index_counter, donor_name,
                            project=None, combined_slab_label=None, slab_count=1):

        # Identifier slab part (unpadded; plural for multi)
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

        # Identifier tile part: "TileN" for numeric, raw text for non-numeric (e.g., "EC")
        tile_part = f"Tile{int(tile)}" if str(tile).isdigit() else tile

        krienen_lab_identifier = (
            f"{current_date}_HMBA_{mit_name}_{slab_part}_{tile_part}_{sort_method}_{modality}{x + 1}"
        )

        experimenter_initials = form_data['sorter_initials'].strip().upper()
        sorting_status = "PS" if sort_method.lower() in ["pooled", "dapi"] else "PN"

        tissue_name = tissue_name_base

        # Project flags and suffixes
        if project == "HMBA_Aim4":
            dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Rseq'
            enriched_prefix = "MPTX"
            rna_suffix = "TX"
        else:
            dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Multiome'
            enriched_prefix = "MPXM"
            rna_suffix = "XR"
        atac_suffix = "XA"

        # Enriched names include initials
        enriched_cell_sample_container_name = f"{enriched_prefix}_{current_date}_{sorting_status}_{experimenter_initials}"
        enriched_cell_sample_name = f'{enriched_prefix}_{current_date}_{sorting_status}_{experimenter_initials}_{port_well}'

        # Study equals Project selection
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

        # Library prep set/name with initials
        key = (library_type, library_prep_date, library_index)
        dup_index_counter[key] = dup_index_counter.get(key, 0) + 1
        library_prep_set = f"{library_type}_{library_prep_date}_{dup_index_counter[key]}"
        library_name = f"{library_prep_set}_{library_index}"

        expected_cell_capture = int(form_data['expected_recovery'])
        concentration = float(form_data['nuclei_concentration'].replace(",", ""))
        volume = float(form_data['nuclei_volume'])
        enriched_cell_sample_quantity_count = round(concentration * volume)

        # Prepare row
        row_data = [
            krienen_lab_identifier,
            seq_portal,
            elab_link,
            current_date,
            mit_name,
            donor_name,
            tissue_name,
            None,  # tissue_name_old
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
            None,  # amplified_cdna_name (filled below for RNA)
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

        # amplified_cdna_name for RNA with initials: AP{INITIALS}{TX|XR}_{date}_{batch}_{A..H}
        if modality == "RNA":
            cdna_amp_date = self.convert_date(form_data['cdna_amp_date'])
            amp_date_key = f"amp_{cdna_amp_date}"
            if amp_date_key not in self.counter_data["amp_counter"]:
                self.counter_data["amp_counter"][amp_date_key] = 0
            reaction_count = self.counter_data["amp_counter"][amp_date_key]
            letter = chr(65 + (reaction_count % 8))
            batch_num = (reaction_count // 8) + 1
            amp_prefix = f"AP{experimenter_initials}{rna_suffix}"
            row_data[21] = f"{amp_prefix}_{cdna_amp_date}_{batch_num}_{letter}"
            self.counter_data["amp_counter"][amp_date_key] += 1

        # Write cells
        for col_num, value in enumerate(row_data, start=1):
            cell = worksheet.cell(row=current_row, column=col_num, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal='left')
            if (modality == "ATAC" and value is None) or (
                modality == "RNA" and col_num == headers.index('ATAC_index') + 1
            ):
                cell.fill = self.black_fill

        # Black fill for tissue_name_old
        tissue_old_col = headers.index('tissue_name_old') + 1
        worksheet.cell(row=current_row, column=tissue_old_col).fill = self.black_fill


# Global instance
data_logger = DataLogger()


@app.route('/')
def index():
    return render_template('index.html')


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
        if success:
            return jsonify({'success': True, 'message': 'Data saved successfully!'})
        else:
            return jsonify({'success': False, 'error': 'Failed to process data'})
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
            # Create a new file for this user if none exists yet
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


@app.route('/get_counter')
def get_counter():
    try:
        return jsonify({
            'next_counter': data_logger.counter_data.get('next_counter', None),
            'success': True
        })
    except Exception as e:
        return jsonify({
            'next_counter': None,
            'success': False,
            'error': str(e)
        })


@app.route('/update_counter', methods=['POST'])
def update_counter():
    try:
        request_data = request.json
        new_counter = request_data.get('new_counter')
        if not isinstance(new_counter, int) or new_counter < 0:
            return jsonify({'success': False, 'error': 'Invalid counter value'})
        data_logger.counter_data['next_counter'] = new_counter
        with open(data_logger.counter_file, 'w') as f:
            json.dump(data_logger.counter_data, f, indent=4)
        return jsonify({'success': True, 'new_counter': new_counter})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


# Optional debug: reset counters (does not rotate per-user pointers)
@app.route('/debug/reset_counter', methods=['POST'])
def debug_reset_counter():
    try:
        data_logger.counter_data["next_counter"] = None
        data_logger.counter_data["date_info"] = {}
        data_logger.counter_data["amp_counter"] = {}
        with open(data_logger.counter_file, 'w') as f:
            json.dump(data_logger.counter_data, f, indent=4)
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    # Local dev server
    app.run(debug=True, host='0.0.0.0', port=8080)