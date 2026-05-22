import hashlib
import json
import os
import sys

import pandas

from loadconfig import getconfig


class CreateBreakDownJson:
    def __init__(self):
        # Preserved even if unused (dead-code detected)
        if getattr(sys, "frozen", False):
            self.app_path = os.path.dirname(sys.executable)
        elif __file__:
            self.app_path = os.path.dirname(__file__)

        self.configFolder, self.breakDownConfig = getconfig()

        self.support_path = os.path.join(
            self.configFolder, "SupportingFiles"
        )

        excel_file = os.path.join(self.support_path, "Breakdown.xlsx")
        json_file = os.path.join(self.support_path, "Breakdown.json")
        md5_hash = os.path.join(self.support_path, "hash_value.txt")
        default_json = os.path.join(
            self.support_path, "defaultValue.json"
        )

        with open(default_json, "r") as default_file:
            self.default_data = json.load(default_file)

        with open(excel_file, "rb") as fh:
            current_md5 = hashlib.md5(fh.read()).hexdigest()

        if os.path.exists(md5_hash):
            with open(md5_hash, "r") as fh:
                old_md5 = fh.read()

            if current_md5 != old_md5:
                self.create_file(json_file)
                with open(md5_hash, "w") as fh:
                    fh.write(current_md5)
        else:
            with open(md5_hash, "w") as fh:
                fh.write(current_md5)
            self.create_file(json_file)

    def create_file(self, json_file):
        json_file_path = os.path.join(
            self.support_path, "Breakdown.json"
        )
        excel_file = os.path.join(
            self.support_path, "Breakdown.xlsx"
        )

        excel_data = pandas.ExcelFile(excel_file)

        journals_details = pandas.read_excel(excel_data, "Journal Names")
        colour_figure_query = pandas.read_excel(
            excel_data, "Colour figure query"
        )
        gq_exact = pandas.read_excel(excel_data, "GQ Exact")
        funding_and_declaration = pandas.read_excel(
            excel_data, "Funding and Declaration"
        )
        supp_exact = pandas.read_excel(excel_data, "SUPP Exact")
        rrh_exact = pandas.read_excel(excel_data, "RRH Exact")
        oa_journals = pandas.read_excel(excel_data, "OA-Journals")

        journal_list = json.loads(
            journals_details.to_json(orient="records")
        )
        oa_list = json.loads(
            oa_journals.to_json(orient="records")
        )
        fundec_list = json.loads(
            funding_and_declaration.to_json(orient="records")
        )
        suppl_list = json.loads(
            supp_exact.to_json(orient="records")
        )
        rrh_list = json.loads(
            rrh_exact.to_json(orient="records")
        )
        cf_list = json.loads(
            colour_figure_query.to_json(orient="records")
        )
        gq_list = json.loads(
            gq_exact.to_json(orient="records")
        )

        first_row = journal_list[0]
        oa_row = oa_list[0]
        fundec_row = fundec_list[0]
        suppl_row = suppl_list[0]

        def norm_name(name):
            return name.replace("&", "A") if name else None

        query_dict = {
            "general_query": {
                3: first_row["General Queries (US)"],
                4: first_row["General Queries (UK)"],
                5: first_row["EmXpert Queries (US)"],
                6: first_row["EmXpert Queries (UK)"],
            },
            "Orchid": first_row["Orchid"],
            "EmXpert Orchid": first_row["EmXpert Orchid"],
            "Funding": first_row["Funding"],
            "EmXpert Funding": first_row["EmXpert Funding"],
            "Funding and Conflict": first_row["Funding and Conflict "],
            "EmXpert Funding and Conflict": first_row[
                "EmXpert Funding and Conflict "
            ],
        }

        funder_dict = {
            "funder": {},
            "declaration": {},
            "end_section_order": {},
        }
        supplementry_dict = {}
        rrh_dict = {}
        cf_dict = {}
        gq_dict = {}
        oa_dict = {}
        journal_dict = {}

        for row in fundec_list:
            jrn_name = norm_name(row["Journal Name"])
            if not jrn_name:
                continue

            funder_dict["funder"].setdefault(jrn_name, {})
            funder_dict["declaration"].setdefault(jrn_name, {})

            loc = row["US/UK"]
            base = "US" if "US" in str(loc) else "UK"

            def pick(col):
                return (
                    row[col]
                    if row[col] is not None
                    else fundec_row[col]
                )

            funder_dict["funder"][jrn_name].update(
                {
                    "AU": pick(f"{base}-Funding (AU)"),
                    "AUs": pick(f"{base}-Funding (AUs)"),
                    "Without_AU": pick(
                        f"{base}-Funding witout  (AU)"
                    ),
                    "Without_AUs": pick(
                        f"{base}-Funding witout  (AUs)"
                    ),
                    "funder_head": row[f"{base}-Funding Head"],
                }
            )

            funder_dict["declaration"][jrn_name].update(
                {
                    "AU": pick(f"{base}-Declaration (AU)"),
                    "AUs": pick(f"{base}-Declaration (AUs)"),
                    "Without_AU": pick(
                        f"{base}-witout Declaration (AU)"
                    ),
                    "Without_AUs": pick(
                        f"{base}-witout Declaration (AUs)"
                    ),
                    "declaration_head": row[
                        f"{base}-Declaration Head"
                    ],
                }
            )

            funder_dict["end_section_order"][
                jrn_name
            ] = row["End section order"]

        for row in suppl_list:
            jrn_name = norm_name(row["Journal Name"])
            if not jrn_name:
                continue

            if row["US/UK"] == "UK":
                supplementry_dict[jrn_name] = suppl_row[
                    "Supplementary details-UK"
                ]
            elif row["US/UK"] == "US":
                supplementry_dict[jrn_name] = suppl_row[
                    "Supplementary details-US"
                ]
            elif row["US/UK"] == "SP":
                supplementry_dict[jrn_name] = (
                    row["Supplementary details-UK"]
                    or row["Supplementary details-US"]
                )
            else:
                supplementry_dict[jrn_name] = None

        for row in rrh_list:
            jrn_name = norm_name(row["Journal Name"])
            if jrn_name:
                rrh_dict[jrn_name] = row[
                    "Running Head details"
                ]

        for row in cf_list:
            jrn_name = norm_name(row["Journal Name"])
            if jrn_name:
                cf_dict[jrn_name] = row[
                    "Colour figure query (US)"
                ]

        for row in gq_list:
            jrn_name = norm_name(row["Journal Name"])
            if jrn_name:
                gq_dict[jrn_name] = row[
                    "General Queries (UK)"
                ]

        for row in oa_list:
            jid = norm_name(row["Journals"])
            if not jid:
                continue

            col = row["column"]
            if col == 3:
                oa_dict[jid] = oa_row[
                    "PDF proofs and SAGE Edit proofs if the CC license has been obtained"
                ]
            elif col == 4:
                oa_dict[jid] = oa_row[
                    "PDF proofs if the CC license has not been obtained"
                ]
            elif col == 5:
                oa_dict[jid] = oa_row[
                    "SAGE Edit proofs if the CC license has not been obtained"
                ]

        for row in journal_list:
            jid = norm_name(row["Journals"])
            if not jid:
                continue

            column = row["column"]
            journal_dict[jid] = {
                "LRH": row["LRH"],
                "BreakDown": False
                if column in (None, "NA")
                else True,
                "HSS/STM": row["HSS/STM"],
                "Article_Type": row["Article Type"],
                "Corresponding_Author": row[
                    "Corresponding Author"
                ],
                "History_Details": row["History details"],
                "Instruction": row["Instruction"],
                "Bio": row["Bio"],
                "FM_Sequence": row["FM Sequence"],
                "BM_Sequence": row["BM Sequence"],
                "funder_text": funder_dict["funder"].get(
                    jid
                ),
                "declaration_text": funder_dict[
                    "declaration"
                ].get(jid),
                "end_section_order": funder_dict[
                    "end_section_order"
                ].get(jid),
                "color_fig_query": cf_dict.get(jid),
                "rrh_format": rrh_dict.get(jid),
                "supplementry_text": supplementry_dict.get(
                    jid
                ),
                "open_access_query": oa_dict.get(jid),
            }

            try:
                journal_dict[jid][
                    "General_Queries"
                ] = query_dict["general_query"][column]
            except Exception:
                journal_dict[jid][
                    "General_Queries"
                ] = gq_dict.get(jid)

            if column in (5, 6):
                journal_dict[jid]["orcid_query"] = query_dict[
                    "EmXpert Orchid"
                ]
                journal_dict[jid]["funding_query"] = query_dict[
                    "EmXpert Funding"
                ]
                journal_dict[jid][
                    "funding_conflict_query"
                ] = query_dict[
                    "EmXpert Funding and Conflict"
                ]
            else:
                journal_dict[jid]["orcid_query"] = query_dict[
                    "Orchid"
                ]
                journal_dict[jid]["funding_query"] = query_dict[
                    "Funding"
                ]
                journal_dict[jid][
                    "funding_conflict_query"
                ] = query_dict[
                    "Funding and Conflict"
                ]

        json_out = {
            "journal_details": journal_dict,
            "default_values": self.default_data,
        }

        with open(json_file_path, "w", encoding="utf-8") as fh:
            json.dump(json_out, fh, indent=4)

        self.remove_substring_from_json(json_file_path)

    def remove_substring_from_json(self, file_path):
        with open(file_path, "r") as fh:
            data = json.load(fh)

        def clean(obj):
            if isinstance(obj, dict):
                return {k: clean(v) for k, v in obj.items()}
            if isinstance(obj, list):
                return [clean(i) for i in obj]
            if isinstance(obj, str):
                return (
                    obj.replace("_x000D_", "")
                    .replace("+T114", "")
                    .replace("+T190", "")
                )
            return obj

        with open(file_path, "w", encoding="utf-8") as fh:
            json.dump(clean(data), fh, indent=4)


# create_json_file = CreateBreakDownJson()
