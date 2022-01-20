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

    # read area_mounting mapper    
    # area mounting -> level
    # area mounting -> drawing grids
    mounting_to_level = {}
    mounting_to_grids = {}
    with open("./config/area_mounting_mapper.csv") as f : 
        r = csv.reader(f)
        for row in r : 
            mounting = row[0].strip()
            level = row[1].strip()
            grids = row[2].strip()
            if mounting not in mounting_to_level : 
                mounting_to_level[mounting] = level
            if mounting not in mounting_to_grids : 
                mounting_to_grids[mounting] = grids

    # read csv data file, and record colmun index
    with open(sys.argv[1]) as f : 
        r = csv.DictReader(f)
        headers = r.fieldnames
        # check REF_DESC and AREA_MOUNTING inside the columns
        if (not "REF_DESC" in headers) or (not "AREA_MOUNTING" in headers) : 
            print("Error : csv does not contains REF_DESC or AREA_MOUNTING.")
            exit()
        # go through each line and append to data pair list
        data_pairs = []
        for line in r : 
            data_pairs.append((line["REF_DESC"].strip(), line["AREA_MOUNTING"].strip()))

        # split by data pair list into different lists based on area mounting
        mounting_to_ref_list = split_area_mounting(data_pairs)

        # get area mounting needs to be generated from user
        # if all area_mountings needs to be generated
        if '-all' in sys.argv[2:] : 
            mountings = list(mounting_to_ref_list.keys())
        else : 
            mountings = sys.argv[2:]

        # loop each area mounting provided by user
        for m in mountings : 
            # if the data pair does not contain the area mounting, skip it
            if not m in mounting_to_ref_list : 
                print("Warning : area mounting {} is not included in data.".format(m))
                continue
            # generate the area mounting
            print("generating area : {}".format(m))
            # prepare the output directory
            output_directory = os.path.join("output" , m)
            prepare_directory(output_directory)
            # loop through all equipment types
            for e in template_mapper : 
                # get the data rows with the correct equipment
                refs = get_refs_by_prefix(mounting_to_ref_list[m], e)
                if not refs : 
                    continue
                # once we got all the data related with this type of equipment
                # further divide the ref list based on area functions (sub system)
                print("\tgenerating equipment type : {}".format(e))
                subs = get_refs_by_subsystem(refs)
                # generate one file for each sub system
                for sub in subs : 
                    filename = os.path.join(output_directory, "L6-{}-{}-{}.xlsx".format(m, e, sub))
                    print("\t\tgenerating subsystem : {}".format(sub))
                    generate_output(sorted(subs[sub]), os.path.join("template", template_mapper[e]), filename, m, mounting_to_level[m], mounting_to_grids[m], sub)

# divide the list of refs based on its subsystem
# receive a list of refs
# returns a dictionary of subsystem -> refs
def get_refs_by_subsystem(refs) : 
    subs = {}
    for e in refs : 
        parts = e.split('.')
        subsys = parts[0]
        if (subsys in subs) : 
            # the sub system is already in the dictionary
            subs[subsys].append(e)
        else : 
            # create a new list of refs
            subs[subsys] = [e]
    return subs

# it receives a list of pairs
# returns a dictionry with area_mounting as key
def split_area_mounting(pairs) : 
    data_by_mounting = {}
    for ref_desc, area_mounting in pairs : 
        # empty strings will be ignored
        if not ref_desc.strip() or not area_mounting.strip(): 
            continue
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




