import sys, os, re, csv, shutil
from lib.excel_handler import generate_output


def main() : 
    # check program parameter
    program_name = os.path.basename(sys.argv[0])
    if len(sys.argv) < 3 : 
        print("usage : {} <filename.csv> <area-mountings>".format(program_name))
        exit()

    # check csv
    file_name = sys.argv[1]
    pattern = r".*\.csv$"
    if (not re.match(pattern, file_name)) : 
        print("usage : {} <filename.csv>".format(program_name))
        exit()

    # read template mapper    
    # prefix -> template file
    template_mapper = {}
    with open("./config/prefix-template-mapper.csv") as f : 
        r = csv.reader(f)
        for row in r : 
            prefix = row[0].strip()
            template = row[1].strip()
            if prefix in template_mapper : 
                print("Error : template mapper duplicated value")
                exit()
            else : 
                template_mapper[prefix] = template

    # read csv file, and record colmun index
    with open(sys.argv[1]) as f : 
        r = csv.DictReader(f)
        headers = r.fieldnames
        # check REF_DESC and AREA_MOUNTING inside the columns
        if (not "REF_DESC" in headers) or (not "AREA_MOUNTING" in headers) : 
            print("Error : csv does not contains REF_DESC or AREA_MOUNTING.")
            exit()
        # go through each line and fill in template accordingly
        data_pairs = []
        for line in r : 
            data_pairs.append((line["REF_DESC"], line["AREA_MOUNTING"]))
        # split by area mounting
        mounting_dict = split_area_mounting(data_pairs)

        # if all area_mountings needs to be generated
        if '-all' in sys.argv[2:] : 
            mountings = list(mounting_dict.keys())
        else : 
            mountings = sys.argv[2:]

        # loop each area mounting
        for m in mountings : 
            if not m in mounting_dict : 
                print("Warning : area mounting {} is not included in data.".format(m))
                continue;
            print("generating area mounting : {}".format(m))
            output_directory = os.path.join("output" , m)
            prepare_directory(output_directory)
            # loop through all equipments
            for e in template_mapper : 
                refs = get_refs_by_prefix(mounting_dict[m], e)
                if not refs : 
                    continue
                filename = os.path.join(output_directory, "L6-{}-{}.xlsx".format(m, e))
                print("generating : {}".format(filename))
                generate_output(refs, os.path.join("template", template_mapper[e]), filename, m)

# it receives a list of pairs
# returns a dictionry with area_mounting as key
def split_area_mounting(pairs) : 
    data_by_mounting = {}
    for ref_desc, area_mounting in pairs : 
        # empty strings will be ignored
        if not ref_desc.strip() or not area_mounting.strip(): 
            continue;
        if area_mounting not in data_by_mounting : 
            data_by_mounting[area_mounting] = [ref_desc]
        else : 
            data_by_mounting[area_mounting].append(ref_desc)
    return data_by_mounting

# it receives a list of ref_descs and a list of prefixes
def get_refs_by_prefix (ref_descs, prefix): 
    rets = []
    for item in ref_descs : 
        if re.match(r".*" + prefix + r"\d{3}$", item) : 
            rets.append(item)
    rets.sort()
    return rets

# prepare an empty directory
def prepare_directory(path) : 
    if os.path.isdir(path) : 
        shutil.rmtree(path)
    os.mkdir(path)

if __name__ == "__main__" : 
    main()




