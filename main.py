from flask import Flask, render_template, request, jsonify, send_file
import os
import json
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import dateutil.parser

app = Flask(__name__)
app.config['SECRET_KEY'] = 'marmoset'


class DataLogger:
    def __init__(self):
        self.config_dir = os.path.join(os.path.expanduser('~'), 'DataLogApp')
        os.makedirs(self.config_dir, exist_ok=True)
        self.counter_file = os.path.join(self.config_dir, 'sample_name_counter.json')
        self.excel_file = os.path.join(self.config_dir, 'krienen_data_log.xlsx')

        self.name_to_code = {
            "Croissant": "CJ23.56.002",
            "Nutmeg": "CJ23.56.003",
            "Jellybean": "CJ24.56.001",
            "Rambo": "CJ24.56.004",
            "Morel": "CJ24.56.015"
        }

        self.tile_location_map = {
            "BRAINSTEM": "BS",
            "BS": "BS",
            "CORTEX": "CX",
            "CX": "CX",
            "CEREBELLUM": "CB",
            "CB": "CB"
        }

        self.black_fill = PatternFill(start_color='000000', fill_type='solid')
        self.load_counter_data()

    def load_counter_data(self):
        if os.path.exists(self.counter_file):
            with open(self.counter_file, 'r') as f:
                try:
                    self.counter_data = json.load(f)
                except json.JSONDecodeError:
                    self.counter_data = {}
        else:
            self.counter_data = {}

        self.counter_data.setdefault("next_counter", 90)
        self.counter_data.setdefault("date_info", {})
        self.counter_data.setdefault("amp_counter", {})

    def convert_date(self, exp_date):
        clean_date = "".join(c for c in exp_date if c.isdigit())
        if len(clean_date) == 6:
            try:
                datetime.strptime(clean_date, '%y%m%d')
                return clean_date
            except ValueError:
                pass
        try:
            parsed_date = dateutil.parser.parse(exp_date)
            return parsed_date.strftime('%y%m%d')
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

    def get_current_time(self):
        return "2025-06-27 17:20:02"

    def get_current_user(self):
        return "lakmecaceres"

    def initialize_excel(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "HMBA"

        headers = ['krienen_lab_identifier', 'seq_portal', 'elab_link', 'experiment_start_date',
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

        ws.append(headers)

        # Apply bold style to headers
        for col_num, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(name="Arial", size=10, bold=True)
            cell.alignment = Alignment(horizontal='left')

        return wb

    def process_form_data(self, form_data):
        # Load or create workbook
        if os.path.exists(self.excel_file):
            workbook = load_workbook(self.excel_file)
        else:
            workbook = self.initialize_excel()

        worksheet = workbook.active

        # Find the last row with content
        last_row_with_content = 1
        for row_idx in range(1, worksheet.max_row + 1):
            row_has_content = any(cell.value is not None for cell in worksheet[row_idx])
            if row_has_content:
                last_row_with_content = row_idx

        current_row = last_row_with_content + 1

        # Process form data
        current_date = self.convert_date(form_data['date'])
        mit_name_input = form_data['marmoset']
        mit_name = "cj" + mit_name_input
        donor_name = self.name_to_code[mit_name_input]

        project = form_data.get('project', '')

        # Process slab and hemisphere
        raw_slab = form_data['slab'].strip()
        hemisphere = form_data['hemisphere'].split()[0].upper()

        if project == "HMBA_CjAtlas_Cortex":
            # Cortex: allow comma-separated slabs, e.g. "9,10,11"
            slab_list = [s.strip() for s in raw_slab.split(',') if s.strip()]
            if not slab_list:
                raise ValueError("No valid slab numbers provided for HMBA_CjAtlas_Cortex")
            combined_slab_label = "_".join(slab_list)
            # Use first slab (zero-padded) where a single numeric slab is needed
            slab = slab_list[0].zfill(2)
        else:
            # Subcortex + Aim 4 + Other: single slab only
            combined_slab_label = None
            slab = raw_slab
            if hemisphere == "RIGHT":
                slab = str(int(slab) + 40).zfill(2)
            elif hemisphere == "BOTH":
                slab = str(int(slab) + 90).zfill(2)
            else:
                slab = slab.zfill(2)

        # Handle tile values
        tile_value = form_data['tile'].strip()
        if tile_value.isdigit():
            tile = str(int(tile_value)).zfill(2)
        else:
            tile = tile_value

        # Process other fields
        tile_location_abbr = form_data['tile_location']
        sort_method = form_data['sort_method']
        sort_method = sort_method.upper() if sort_method.lower() == "dapi" else sort_method

        # FACS population
        if sort_method.lower() == "pooled":
            facs_population = form_data['facs_population']
        elif sort_method.lower() == "unsorted":
            facs_population = "no_FACS"
        else:
            facs_population = "DAPI"

        rxn_number = int(form_data['rxn_number'])

        # Update counters
        if current_date not in self.counter_data["date_info"]:
            self.counter_data["date_info"][current_date] = {
                "total_reactions": 0,
                "batches": []
            }

        date_info = self.counter_data["date_info"]
        date_entry = date_info[current_date]
        existing_total = date_entry["total_reactions"]

        total_reactions_after = existing_total + rxn_number
        batches_before = (existing_total + 7) // 8
        batches_after = (total_reactions_after + 7) // 8
        new_batches_needed = batches_after - batches_before

        new_p_numbers = [self.counter_data["next_counter"] + i for i in range(new_batches_needed)]
        self.counter_data["next_counter"] += new_batches_needed

        all_batches = date_entry["batches"].copy()
        all_batches.extend({"p_number": p, "count": 0} for p in new_p_numbers)

        # Calculate port wells
        port_wells = []
        for x in range(rxn_number):
            global_idx = existing_total + x + 1
            batch_idx = (global_idx - 1) // 8
            p_number = all_batches[batch_idx]["p_number"]
            port_well = (global_idx - 1) % 8 + 1
            port_wells.append((p_number, port_well))

        # Update counters
        date_entry["total_reactions"] = total_reactions_after
        date_entry["batches"] = all_batches

        # Process indices
        atac_indices = [self.convert_index(index) for index in form_data['atac_indices'].split(",")] if form_data.get('atac_indices') else []
        atac_indices = [self.pad_index(index) for index in atac_indices]

        rna_indices = [self.convert_index(index) for index in form_data['rna_indices'].split(",")] if form_data.get('rna_indices') else []
        rna_indices = [self.pad_index(index) for index in rna_indices]

        # Process the data for each reaction and modality
        dup_index_counter = {}
        headers = [cell.value for cell in worksheet[1]]

        # Modalities depend on project:
        # - Aim 4: RNA only (no ATAC)
        # - All others (Subcortex, Cortex, Other): RNA + ATAC
        if project == "Aim 4":
            modalities = ["RNA"]
        else:
            modalities = ["RNA", "ATAC"]

        for x in range(rxn_number):
            p_number, port_well = port_wells[x]
            barcoded_cell_sample_name = f'P{str(p_number).zfill(4)}_{port_well}'

            tissue_name_base = f"{donor_name}.{tile_location_abbr}.{slab}.{tile}"

            for modality in modalities:
                self.write_modality_data(
                    worksheet, current_row, modality, x,
                    current_date, mit_name, slab, tile, sort_method,
                    port_well, barcoded_cell_sample_name,
                    form_data,
                    tissue_name_base=tissue_name_base,
                    rna_indices=rna_indices, atac_indices=atac_indices,
                    headers=headers, dup_index_counter=dup_index_counter,
                    donor_name=donor_name,
                    project=project,
                    combined_slab_label=combined_slab_label
                )
                current_row += 1

        # Save workbook and counter data
        workbook.save(self.excel_file)
        with open(self.counter_file, 'w') as f:
            json.dump(self.counter_data, f, indent=4)

        return True

    def write_modality_data(self, worksheet, current_row, modality, x, current_date, mit_name, slab, tile, sort_method,
                            port_well, barcoded_cell_sample_name, form_data, tissue_name_base, rna_indices,
                            atac_indices, headers, dup_index_counter, donor_name,
                            project=None, combined_slab_label=None):

        # Create krienen_lab_identifier
        if project == "HMBA_CjAtlas_Cortex" and combined_slab_label:
            # Example: 251118_HMBA_cjMorel_Slabs_9_10_11_Tile22_...
            slab_part = f"Slabs_{combined_slab_label}"
        else:
            # Original behavior: single slab with numeric value
            slab_part = f"Slab{int(slab)}"

        tile_part = f"Tile{int(tile) if tile.isdigit() else tile}"

        krienen_lab_identifier = (
            f"{current_date}_HMBA_{mit_name}_{slab_part}_{tile_part}_{sort_method}_{modality}{x + 1}"
        )

        sorter_initials = form_data['sorter_initials'].strip().upper()
        sorting_status = "PS" if sort_method.lower() in ["pooled", "dapi"] else "PN"

        tissue_name = tissue_name_base
        dissociated_cell_sample_name = f'{current_date}_{tissue_name}.Multiome'
        enriched_cell_sample_container_name = f"MPXM_{current_date}_{sorting_status}_{sorter_initials}"
        enriched_cell_sample_name = f'MPXM_{current_date}_{sorting_status}_{sorter_initials}_{port_well}'

        study = "HMBA_CjAtlas_Subcortex" if form_data['project'] == "HMBA_CjAtlas_Subcortex" else form_data.get(
            'project_name', '')

        seq_portal = "no"
        elab_link = form_data.get('elab_link', '')
        facs_population = form_data.get('facs_population', 'no_FACS')
        cell_prep_type = "nuclei"

        library_prep_date = (self.convert_date(form_data['rna_prep_date']) if modality == "RNA"
                             else self.convert_date(form_data['atac_prep_date']))

        if modality == "RNA":
            library_method = "10xMultiome-RSeq"
            library_type = "LPLCXR"
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
        else:  # ATAC
            library_method = "10xMultiome-ASeq"
            library_type = "LPLCXA"
            library_index = atac_indices[x]

            atac_concentration = float(form_data['atac_lib_concentration'].split(',')[x])
            lib_quant = atac_concentration * 20

            atac_size = int(form_data['atac_sizes'].split(',')[x])
            library_cycles = int(form_data['library_cycles_atac'].split(',')[x])

            # RNA-only values not used for ATAC
            cdna_concentration = None
            cdna_amplified_quantity = None
            cdna_library_input = None
            percent_cdna_400bp = None
            cdna_pcr_cycles = None
            rna_size = None

        # Update library prep set counter
        key = (library_type, library_prep_date, library_index)
        dup_index_counter[key] = dup_index_counter.get(key, 0) + 1
        library_prep_set = f"{library_type}_{library_prep_date}_{dup_index_counter[key]}"
        library_name = f"{library_prep_set}_{library_index}"

        expected_cell_capture = int(form_data['expected_recovery'])
        concentration = float(form_data['nuclei_concentration'].replace(",", ""))
        volume = float(form_data['nuclei_volume'])
        enriched_cell_sample_quantity_count = round(concentration * volume)

        # Prepare row data
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
            "10xMultiome-RSeq" if modality == "RNA" else None,
            self.convert_date(form_data['cdna_amp_date']) if modality == "RNA" else None,
            None,  # amplified_cdna_name (filled later for RNA)
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

        # Handle amplified_cdna_name for RNA
        if modality == "RNA":
            cdna_amp_date = self.convert_date(form_data['cdna_amp_date'])
            amp_date_key = f"amp_{cdna_amp_date}"

            if amp_date_key not in self.counter_data["amp_counter"]:
                self.counter_data["amp_counter"][amp_date_key] = 0

            reaction_count = self.counter_data["amp_counter"][amp_date_key]
            letter = chr(65 + (reaction_count % 8))
            batch_num_for_amp = (reaction_count // 8) + 1

            row_data[21] = f"APLCXR_{cdna_amp_date}_{batch_num_for_amp}_{letter}"
            self.counter_data["amp_counter"][amp_date_key] += 1

        # Write to Excel
        for col_num, value in enumerate(row_data, start=1):
            cell = worksheet.cell(row=current_row, column=col_num, value=value)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal='left')

            # Apply black fill for certain cells
            if (modality == "ATAC" and value is None) or (
                    modality == "RNA" and col_num == headers.index('ATAC_index') + 1):
                cell.fill = self.black_fill

        # Apply black fill to tissue_name_old
        tissue_old_col = headers.index('tissue_name_old') + 1
        worksheet.cell(row=current_row, column=tissue_old_col).fill = self.black_fill


# Create global instance
data_logger = DataLogger()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/submit', methods=['POST'])
def submit_data():
    try:
        form_data = request.json

        # Validate required fields
        required_fields = ['date', 'marmoset', 'slab', 'tile', 'hemisphere', 'tile_location', 'sort_method',
                           'rxn_number']
        for field in required_fields:
            if not form_data.get(field):
                return jsonify({'success': False, 'error': f'Missing required field: {field}'})

        # Process the data
        success = data_logger.process_form_data(form_data)

        if success:
            return jsonify({'success': True, 'message': 'Data saved successfully!'})
        else:
            return jsonify({'success': False, 'error': 'Failed to process data'})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/download')
def download_excel():
    if os.path.exists(data_logger.excel_file):
        return send_file(data_logger.excel_file, as_attachment=True, download_name='krienen_data_log.xlsx')
    else:
        return jsonify({'error': 'No data file found'}), 404


@app.route('/get_counter')
def get_counter():
    try:
        return jsonify({
            'next_counter': data_logger.counter_data.get('next_counter', 90),
            'success': True
        })
    except Exception as e:
        return jsonify({'next_counter': 90, 'success': False, 'error': str(e)})


@app.route('/update_counter', methods=['POST'])
def update_counter():
    try:
        request_data = request.json
        new_counter = request_data.get('new_counter')

        if not isinstance(new_counter, int) or new_counter < 0:
            return jsonify({'success': False, 'error': 'Invalid counter value'})

        # Update the counter in the data_logger instance
        data_logger.counter_data['next_counter'] = new_counter

        # Save to file
        with open(data_logger.counter_file, 'w') as f:
            json.dump(data_logger.counter_data, f, indent=4)

        return jsonify({'success': True, 'new_counter': new_counter})

    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)