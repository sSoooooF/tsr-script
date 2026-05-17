import csv
import io
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from pathlib import Path

TSR_DIR = Path(__file__).resolve().parent
OUTPUT_CSV = TSR_DIR / 'dell_tsr_audit.csv'
OUTPUT_MD = TSR_DIR / 'dell_tsr_audit.md'

COMPONENT_CLASSES = {
    'DCIM_CPUView': 'CPU',
    'DCIM_MemoryView': 'DIMM',
    'DCIM_PowerSupplyView': 'PS',
    'DCIM_PhysicalDiskView': 'Drive',
    'DCIM_PCIDeviceView': 'PCI',
}

VIEW_PATH = 'tsr/hardware/sysinfo/inventory/sysinfo_DCIM_View.xml'


def get_value(prop):
    if prop is None:
        return ''
    value = prop.find('VALUE')
    if value is not None and value.text is not None:
        return value.text.strip()
    # fallback to DisplayValue if VALUE is absent or empty
    display = prop.find('DisplayValue')
    return display.text.strip() if display is not None and display.text else ''


def instance_props(inst):
    props = {}
    for prop in inst.findall('PROPERTY'):
        name = prop.attrib.get('NAME')
        if not name:
            continue
        props[name] = get_value(prop)
    return props


def open_nested_tsr(outer_zip_path: Path):
    with zipfile.ZipFile(outer_zip_path, 'r') as outer:
        nested_paths = [name for name in outer.namelist() if name.lower().endswith('.pl.zip')]
        if not nested_paths:
            raise FileNotFoundError(f'No nested .pl.zip found in {outer_zip_path}')
        with outer.open(nested_paths[0]) as nested_file:
            nested_bytes = nested_file.read()
        return zipfile.ZipFile(io.BytesIO(nested_bytes), 'r')


def parse_tsr_archive(outer_zip_path: Path):
    with open_nested_tsr(outer_zip_path) as nested:
        if VIEW_PATH not in nested.namelist():
            raise FileNotFoundError(f'{VIEW_PATH} not found inside {outer_zip_path}')
        text = nested.read(VIEW_PATH)
    root = ET.fromstring(text)

    system_instances = [inst for inst in root.findall('.//INSTANCE') if inst.attrib.get('CLASSNAME') == 'DCIM_SystemView']
    server = instance_props(system_instances[0]) if system_instances else {}

    server_model = server.get('Model', 'Unknown')
    server_serial = server.get('ServiceTag', server.get('ChassisServiceTag', 'Unknown'))
    server_part = server.get('BoardPartNumber', server.get('PartNumber', 'Unknown'))
    server_tag = server.get('NodeID', server_serial)

    # Collect component instances by class
    components = defaultdict(list)
    for inst in root.findall('.//INSTANCE'):
        cname = inst.attrib.get('CLASSNAME')
        if cname in COMPONENT_CLASSES:
            comp = instance_props(inst)
            comp['__class'] = cname
            components[cname].append(comp)

    rows = []

    def add_grouped(cname, rows_by_key):
        comp_type = COMPONENT_CLASSES[cname]
        for key, items in rows_by_key.items():
            count = len(items)
            sample = items[0]
            part = sample.get('PartNumber') or sample.get('PPID') or sample.get('FQDD') or sample.get('ProcessorID') or sample.get('SerialNumber') or 'N/A'
            if cname == 'DCIM_PhysicalDiskView':
                part = sample.get('PartNumber') or sample.get('PPID') or sample.get('SerialNumber') or sample.get('FQDD') or 'N/A'
            if cname == 'DCIM_PowerSupplyView':
                part = sample.get('PartNumber') or sample.get('PPID') or sample.get('FQDD') or 'N/A'
            if cname == 'DCIM_CPUView':
                part = sample.get('PartNumber') or sample.get('ProcessorID') or sample.get('FQDD') or sample.get('SerialNumber') or 'N/A'
            descriptions = {item.get('Description') or item.get('DeviceDescription') or item.get('Model') or item.get('SocketDesignation') or '' for item in items}
            desc = '; '.join(sorted(d for d in descriptions if d))
            if not desc:
                desc = key
            rows.append({
                'server_name': server_model,
                'server_serial': server_serial,
                'server_part': server_part,
                'component': comp_type,
                'quantity': count,
                'component_part': part,
                'description': desc,
            })

    # grouping rules
    if 'DCIM_CPUView' in components:
        cpu_items = components['DCIM_CPUView']
        key = lambda item: (item.get('PartNumber') or item.get('ProcessorID') or item.get('Model') or item.get('SocketDesignation') or item.get('Description') or 'CPU')
        grouped = defaultdict(list)
        for item in cpu_items:
            grouped[key(item)].append(item)
        add_grouped('DCIM_CPUView', grouped)

    if 'DCIM_MemoryView' in components:
        mem_items = components['DCIM_MemoryView']
        key = lambda item: (item.get('PartNumber') or item.get('DeviceLocator') or item.get('Model') or item.get('SerialNumber') or 'DIMM')
        grouped = defaultdict(list)
        for item in mem_items:
            grouped[key(item)].append(item)
        add_grouped('DCIM_MemoryView', grouped)

    if 'DCIM_PowerSupplyView' in components:
        ps_items = components['DCIM_PowerSupplyView']
        key = lambda item: (item.get('PartNumber') or item.get('FQDD') or item.get('DeviceDescription') or 'PS')
        grouped = defaultdict(list)
        for item in ps_items:
            grouped[key(item)].append(item)
        add_grouped('DCIM_PowerSupplyView', grouped)

    if 'DCIM_PhysicalDiskView' in components:
        disk_items = components['DCIM_PhysicalDiskView']
        key = lambda item: (item.get('PPID') or item.get('PartNumber') or item.get('Model') or item.get('DeviceDescription') or item.get('FQDD') or 'Drive')
        grouped = defaultdict(list)
        for item in disk_items:
            grouped[key(item)].append(item)
        add_grouped('DCIM_PhysicalDiskView', grouped)

    if 'DCIM_PCIDeviceView' in components:
        pci_items = components['DCIM_PCIDeviceView']
        key = lambda item: (item.get('DeviceDescription') or item.get('Description') or item.get('FQDD') or item.get('Model') or 'PCI')
        grouped = defaultdict(list)
        for item in pci_items:
            grouped[key(item)].append(item)
        add_grouped('DCIM_PCIDeviceView', grouped)

    return server_tag, rows


def main():
    zip_dir = TSR_DIR
    outer_zips = sorted(zip_dir.glob('*.zip'))
    all_rows = []
    idx = 1
    if not outer_zips:
        raise FileNotFoundError('No ZIP archives found in directory')

    for outer in outer_zips:
        try:
            server_tag, rows = parse_tsr_archive(outer)
        except Exception as exc:
            print(f'Failed parse {outer.name}: {exc}')
            continue
        for row in rows:
            row['index'] = idx
            idx += 1
            row['outer_file'] = outer.name
            all_rows.append(row)

    fieldnames = ['index', 'outer_file', 'server_name', 'server_serial', 'server_part', 'component', 'quantity', 'component_part', 'description']
    with OUTPUT_CSV.open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    with OUTPUT_MD.open('w', encoding='utf-8') as f:
        f.write('| index | archive | server name | server serial | server part | component | quantity | component part | description |\n')
        f.write('|---|---|---|---|---|---|---|---|---|\n')
        for row in all_rows:
            cells = [str(row[col]).replace('|', '\|') for col in fieldnames]
            f.write('| ' + ' | '.join(cells) + ' |\n')

    print(f'Parsed {len(all_rows)} component rows from {len(outer_zips)} archives')
    print(f'CSV output: {OUTPUT_CSV}')
    print(f'Markdown output: {OUTPUT_MD}')


if __name__ == '__main__':
    main()
