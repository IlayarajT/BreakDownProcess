import re
import spacy
import warnings
# from spacy.language import Language


# @Language.factory("curated_transformer")

class PyLabelAuthor:
    def __init__(self):
        self.file_type = ''
        self.customer = ''
        self.hyper_text = {'*': 'lowast', '#': 'hash', '|': 'par', '||': 'dpar', '†': 'dagger', '‡': 'ddagger',
                      '§': 'section', '¶': 'para', '$': 'dollar', '**': 'lowastlowast', '***': 'lowastlowastlowast',
                      '****': 'lowastlowastlowastlowast', '*****': 'lowastlowastlowastlowastlowast'}
        self.degree = r"ScM, CGC|MSW, LCSW|MD, FANPA|PharmD,\s*CGP|MPsych\(Hons\)\,\s*PhD\,\s*FASSA|PhD, MD, FRANZCP|MBBS, " \
                 "FRANZCP|MBBS BMedSci FRACS FAOrthA|Ph.D.,? ?BCBA-D ?(Chair)?|M.B.A.?,? ?|AGSF|ALM|A.?S.?C.?P.?|AB|" \
                 "ACS-OB/GYN|ACS-RA|AM|ANP|APN-BC|APRN-BC|APRN|ART|AuD|BAcc|BAcy|BAE|BAI|BAppSc|BArch|BASc?|BASET BE|BBA" \
                 "|BCD (Hon)|B.?A.?|BChD?|BChiro?|BCL?|BCoun|B.?D.?S.?(Singapore)|BDent|BDes|BDiv|BDSc?|BD|BEcon&Fin|BEcon" \
                 "|BEng|BEd, GradDipHealthEd|BEd|BE|BFin|BHSc|BLA LArch|BLitt|BME|BMedSc|BMed|BMid|BMSc|BMin|BMus|BM BCh|" \
                 "BM BS|BMBCh|BMBS|B.?M.?L.?S.?|BNat|BNurs|BN|BOH|BOptom|BOstMed|BOst|BPhil|BPL|BPod|BPharm|B.?S.?c.?,? ?" \
                 "|BSc (Econ)|BSc (Eng)|BSc (Osteo)|BSc (Psych)|BScEng|BScN|BSc-PharmSci|BScPhm|BSDH|BSET|BSE|BSF|BSN" \
                 "|BSocSc|B.?Sc.? (Vet)|B.?Sc.?|B.?S.?|BSN|BSPT|BSRC|BSW|BTchg|BTheol|BTh|BURPl|B.?Vet.?Med.?|BVM&S" \
                 "|B.?V.?M.?S.? (Hons)|B.?V.?Sc.?|B.?V.?Sc.?|B.?Tech.?|B.?Vet.?Sc.?|BVetMed MRCVS|BSRT(R)(N)|BVS|NCT|CB" \
                 "|CD|ChB|ChD|C-FNP|ChM|CHA|CNS|CCDS|CHES|CIBS|CIH|CMBS-I|CMC|CMD|CDCES|CMNT|CNMT|CNM|CNOR|COMT|CO|CPC-H" \
                 "|CPC-I|CPC|CPFT|CPI|CPNP|CRNA|CRNFA|CRNP|CRTT|CRT|CTR|CertLAS|CertZooMed|DArch|DBA|DBS|DCL|DCh|DC" \
                 "|DClinDent|DClinPsych|DCM|DDSc|D.?D.?S.?|DD|DEd|DHSci?|DHL|DipArch|Dip.?Ag.?Sci.?|Dip VCS|Dip.?Orth" \
                 "|Dipl.? A.?C.?V.?P.?|Dipl.? A.?C.?Z.?M.?|Dipl.? AVCO|Dipl.? ECAMS|Diplomate ACVIM|DLitt|DMin|DMSc|DMus" \
                 "|DMD|DME|DMSc|DMS|DMV|DM|DMRD|DNAP|DNP|DNB|DNE|DNS|DNSc|D.?O.?|DPH|DN|DNursSci|DOM|DO|DPharm|D ?Phil" \
                 "|DPhysio|DProf|DPs|DPT|DrMedDent|DrMed|DrMuD|DrOdont|DPM|DP|DrMedVet|DrPH|DrOT|DrPH|DSocSci|DSW|DS" \
                 "|DThP?|DUniv|DTM&H|DTPH|DVetMed|DVMS|D.?V.?M.?|DVSc|DVS|Dr.? ?Med.? ?Vet.?|Dr.? Med.? Dent|Dr.?Odont.?" \
                 "|EdB|Ed.?D|EdPsychD|EngD|HScD|JD,? NA|JD|JSD|LittB|LittD|LLB|LLD|LLM|ELS|EMT-P|EMT|FAAAI|FAAFP|FAAN" \
                 "|FAAOS|FAAO|FAAP|FACCP|FACC|FACEP|FACE|FACOFP|FACOG|FACP|FACSM|FACS|F.?A.?M.?S.?|FAHA|FAMS (Singapore)" \
                 "|FAOTA|FAPHA|FAPTA|FCGP|FCIPD|FCPS|FCAHS|FCS|F.?D.?S.?R.?C.?S.?Ed.?|FFARCS|FFA|FFSEM|FNP-BC|FNP-BC|FNP" \
                 "|FNASN|FRACP|FRCA|FRCGP|FRCOG|FRCP(C)|FRCP(Edin)|FRCP(Edinburgh)|FRCP(Glasg)|FRCP(Glasgow)|FRCP(Ire)" \
                 "|FRCP(Ireland)|FRCPath|FRCPC|FRCPE|FRCPI|FRCP|FRCR|FRCS(C)|FRCS(Edin)|FRCS(Edinburgh)|FRCS(Glasg)" \
                 "|FRCS(Glasgow)|FRCS(Ire)|FRCS(Ireland)|FRCSC|FRCSEd|FRCSE|FRCSI|FRCS|FRCVS|FRS|FSNMTS|FRCOphth|GNP" \
                 "|GSAF|HSD|JD|LLB|LLD|LLM|LPN|LVN|M(ASCP)|MAcc|MAcy|MAc|MAE|MAOT|MArch|MASc|MASP|MAPP|MASM|M.?A.?" \
                 "|MB BChir|MB BCh|MB ChB|MB,BChir|MB,BCh|M.?B.?B.?S.?|MB,ChirB|MB,ChB|MBA|MBBChir|MBBCh|MBBS|MBChB" \
                 "|MBiochem|MBiolSci|MBiol|M.Biostat|MBM|M.?B.?|MChem|MChiro?|MClinDent|MCMC|MComp|MCP|MCh" \
                 r"|M.?D.?S.?(Orthodontics)|MD,\s*MsC|MD,\s*PhD|MD-PhD|MDent|MDiv|MDrama|MDCM|M.?D.?|M.?Ed.?|MDSc?" \
                 "|MEarthSci|MEcon|MEng|MEnvSci|MESci|MES|MFA|MF|MGeog|MGeol|MGeophys|MHA|MInf|MJur|MLA LArch|MLib" \
                 "|MLIS|MLitt|MMathComp|MMathPhys|MMathStat|MMath|MMin|MMORSE|MMSc|MMus|MIHM|M.?L.?S.?" \
                 "|MMed.? Int. Med. (Singapore)|MMM|MNatSci?|MNursSci|MN|MOcean|MOT|MPAff|M.?P.?A.?|M.?P.?H.?|MPharm" \
                 "|M.?Phil.?|MPhys|MPlan|MPl|MPPA|MPP|MPT|MProfSt|MPS|MRes|MRP|MRCP(UK)|MRCP|MRCSI|MRCS|M.? ?Sc.?" \
                 "|MSAcy|MScAg|MSIM|MSIS|MSN|MSPH|MStat|MSurg|M.?S.?D.?|MSW|M.?S.?|MT(ASCP)|MSOM|MSOT|MSPAS|MSPH|MSPT" \
                 "|MSSc|MSocSc|MStat|MSt|MSTOM|MSUP|MSurg|MSW|MTheol|MTh|MTL|MTA|MTP|MT|MUEP|MUniv|MUS|MUP|MusB|MusD" \
                 "|MusM|MVB|MVS|MSCI?|Msci?|MSci?|M.?T.?|M.?Orth.?Ed.?|NCSN|ND|NP|NMD|OD|OTD?|OTR|PA(?:-C)?|PD|PG Dip" \
                 "|PharmD|PhD(c)|PH.?D.?|P.?h.?D.?|Ph.?D|PhG|Ph|PMHCNS-BC|PNP-BC|PNP|Prof.? Dr.? ?Med.? ?Vet.?|PCS|PodB" \
                 "|PodD|PrD|PsyD|PT|RDH?|RN,C|R.?N.?|RNA|RNC|RPFT|RPh|RPT|RRL|RTR|RT|RRT|ScD|SJD|SocSciD|STB|STD|STL|STM" \
                 "|S.?M.? (ASCP)|SM|ThB|ThD|ThM|VetMB|VMD|(Hons)|DACVIM|WHNP|FHRS|GradDip (Clin Epi)|BOrth (Hons)|CAISS" \
                 r"|BPhysio(Hons)|Grad Dip Biostats|MAppSc|FACEM|FRACS|MEHP|MHPE|MHS|M.?Med|BA(Hons),\s*PhD" \
                 "|BSc(Hons), BA, PhD|BSc(Hons).?,? ?|MRCS.?,? ?|MRes.?,? ?|MD.?,? ?|FRCS(Orth).?,? ?" \
                 "|FRCS(Trs&sOrth).?,? ?|DPhil.?,? ?|L.I.S.W.?,? ?|M.ed.?,? ?|M.D.?,? ?|PhD.?,? ?|FRCSI.?,? ?" \
                 "|FRCSEd(Orth).?,? ?|FAMS(Orth).?,? ?|MBBS(Sing).?,? ?|PhD(Math).?,? ?|MBBS.?,? ?|MMed.?,? ?" \
                 "|FRCS.?,? ?|MChOrtho.?,? ?|FAMS.?,? ?|MPH&TM.?,? ?|MPH|MMed(Ortho).?,? ?|Dip SpMed.?,? ?" \
                 "|MMed(Surg).?,? ?|FRCS(T&O).?,? ?|MSN.?,? ?|ACNP.?,? ?|ANP.?,? ?|FAANP.?,? ?|PA-C.?,? ?" \
                 "|MSc.?,? ?|FRCS(C).?,? ?|PharmD.?,? ?|DScs?.?,? ?|MPS.?,? ?|BPharms(Hons).?,? ?|Ph.D.?,? ?" \
                 r"|B.Ch.?,? ?|F.A.S.M.B.S.?,? ?|C.P.E.?,? ?|F.R.C.S.?,? ?|IBHRE|BM|PMHNP-BC\,? ?|NEA-BC|PGDEPI" \
                 "|Dr. PH|M.R.C.S|CCT|MMedSc|MMED|BHB|FRCPA|FRCPCH|FFSc|ABMLI|FRCPATH|CNE|SCFES|Bsc"
        self.AffSymbolDesigs = r'(?:(?:&amp;|&par;|\&AFFMLIpar|\&AFFMLpar|\&AFFMLIBar|\&AFFML2BAR|\&AFFML2BAR|\&AFFMLIpar' \
                          r'|\&AFFMLpar|\&AFFMLIBar|\&Verbar\;|\&ASH|\&Dagger\;|\&dagger\;|\&sect\;|\&para\;|\&xutri\;' \
                          r'|\&boxV\;|\&\#9651\;|\&\#8224\;|\&\#8225\;|\&\#12289\;|\&\#449\;|\&\#167\;|\&\#450\;' \
                          r'|\&\#8710\;|\&hearts\;|(?:[\\*†$‡§¶@\^\\#\\|\\+]+))+)'
        self.role = r'Associate Professors?|Candidate|Prof\.|(?:Assistant )?Professor' \
               r'|Editors?\-in\-Chief|Dr Med|Dr Odont|(?:Guest )?Editors?|Associate Editors?' \
               '|(?:Graduate )?Research Assistant|Moderator|Student|Managing Editors?'
        self.gen = r'Jr\.?|Sr\.?|I{2,3}|I[VX]|V[I]+'
        self.pronouns = '(she(, )?|her(, )?|he(, )?|him(, )?|they(, )?|them(, )?|in conversation with)+'
        self.AuthorPrefix = '(?:' + '|'.join(
            [r'[A|E]l ', 'Ap ', 'Ben ', r'Dell(?:[a|e])? ', 'Dalle ', r'D[a|e]ll\'', 'Dela ', 'Del ',
             r'(?:De|de) (?:La |Los )?', r'[dD][a|i|u] ', '[lL][a|e|o] ', r'[D|L|O]\'', r'St\.? ', 'San ', 'Den ',
             r'[vV]on (?:Der )?', r'[vV]an (?:[dD]e(?:n|r)? )?']) + ')'

    def author_process(self, author):
        TaggedAuthors = ""
        commentid = ""
        comment = ""
        new_degree, person_list = self.predict_degrees(author, self.degree)
        new_degrees = "|".join(new_degree)
        self.degree = new_degrees + "|" + self.degree
        while re.search(r'(<querycomment>[A-Z]{1,2}[0-9]+\:</querycomment>)', author, re.IGNORECASE):
            comment += re.search(r'(<querycomment>[A-Z]{1,2}[0-9]+\:</querycomment>)', author, re.IGNORECASE).group(1)
            author = re.sub(r'(<querycomment>[A-Z]{1,2}[0-9]+\:</querycomment>)', '', author, re.IGNORECASE)
        author = re.sub(r'<hyperlink\s[^>]*>(\(?http:\/\/orcid\.org/.*?)</hyperlink>', r'\1', author,
                        flags=re.IGNORECASE)
        if re.search(r'<br>', author):
            author = re.sub(r'([, |; ]*)<br>(?:</br>)?', r'</au><x>\1</x><au>', author)
        author = re.sub(r'<\/?(b|i|u|scp|caps|strong)([0-9])?>', '', author, flags=re.IGNORECASE)
        author = re.sub(r'</?acepara>', '', author, flags=re.IGNORECASE) if re.search(r'^(cucenter)$', self.file_type,
                                                                                      re.IGNORECASE) else author
        author = re.sub(r'\,([A-Z])', r', \1', author)
        author = re.sub(r'(\()\s*((?:(?:&\#(?:[0-9]+)\;)+))\s((?:(?:&\#(?:[0-9]+)\;)+))\s*(\))', r'\1\2\3\4', author,
                        flags=re.IGNORECASE)
        author = re.sub(r'([0-9]+|[a-z])([^0-9a-z]</sup>)', r'\1,\2', author, flags=re.IGNORECASE)
        author = re.sub(r'(List of FACE-BD collaborators|the NELA study group)(\s*[;,.]\s*|\*?<sup>|\*?</p>)',
                        r'<collab>\1</collab>\2', author, flags=re.IGNORECASE)
        author = re.sub(r'([a-z0-9])([*]+<\/sup>)', r'\1,\2', author, flags=re.IGNORECASE)
        author = re.sub(r'(<sup>)([a-z])([a-z])(</sup>)', r'\1\2,\3\4', author, flags=re.IGNORECASE)
        if len(person_list) > 0:
            for person in person_list:
                author = re.sub(rf'({person})', r'<persion>\1</persion>', author, flags=re.IGNORECASE)
        author = re.sub(rf'(\,|\s)({self.degree})(?=([^a-zA-Z]|$))', r'\1<deg>\2</deg>\3', author,
                        flags=re.IGNORECASE) if not re.search(r'<deg>', author) else author
        author = re.sub(r'<</deg>#</deg><', r'</deg><', author, flags=re.IGNORECASE)
        author = re.sub(r'(</deg>)(\,?\s?\([^\)0-9]+\))', r'\2\1', author, flags=re.IGNORECASE)
        author = re.sub(r'(</deg>)(\&amp;TM)', r'\2\1', author, flags=re.IGNORECASE)
        author = re.sub(r'</deg>(\s[\(\w+\)\,]+\s)<deg>', r'\1', author, flags=re.IGNORECASE)
        author = re.sub(r'</deg>([\,\s]+)<deg>', r'\1', author, flags=re.IGNORECASE)
        author = re.sub(r'<deg>(\s*)</deg>', r'\1', author, flags=re.IGNORECASE)
        author = re.sub(r'<\/deg>\/L', '/L</deg>', author, flags=re.IGNORECASE)
        author = re.sub(r'(\; |\;|\, |\,)</deg>', r'</deg>\1', author, flags=re.IGNORECASE)
        author = re.sub(r'(\sand\s)<deg>([^><]+)</deg>', r'\1\2', author, flags=re.IGNORECASE)
        author = re.sub(r'(\;|\.|\,|)(\sand\s)', r'<x>\1\2</x>', author, flags=re.IGNORECASE)
        author = re.sub(r'(<x>(((?!</?x>).)*)<\/x>)<\/deg>', r'</deg>\1', author)
        author = re.sub(r'(<sup[^>]*>)(\*)(\&[^\;]+;)(</sup>)', r'\1\2\4\1\3\4', author, flags=re.IGNORECASE)
        author = re.sub(r'(&Dagger;)', '‡', author)
        author = re.sub(r'(&dagger;)', '†', author)
        author = re.sub(r'(&sect;)', '§', author, flags=re.IGNORECASE)
        author = re.sub(r'(&para;)', '¶', author, flags=re.IGNORECASE)
        author = re.sub(r'(&hash;)', '#', author, flags=re.IGNORECASE)
        author = re.sub(r'(&par;)', '|', author, flags=re.IGNORECASE)
        author = re.sub(r'(&amp;)', '&', author, flags=re.IGNORECASE)
        author = re.sub(r'(&lowast;)', '*', author, flags=re.IGNORECASE)
        author = re.sub(r'(&amp;&amp;)', '&&', author, flags=re.IGNORECASE)
        author = re.sub(r'(&amp;&amp;&amp;)', '&&&', author, flags=re.IGNORECASE)
        author = re.sub(r'(&amp;&amp;&amp;&amp;)', '&&&&', author, flags=re.IGNORECASE)
        author = re.sub(r'(&ndash;)', '–', author, flags=re.IGNORECASE)
        author = re.sub(r'\.</p>', r'<x>.</x></p>', author)
        author = re.sub(r'[\(\[]corresponding\sauthor[\)\]]', r'<sup>*</sup>', author)
        author = re.sub(r'</snm>,<sup>', r'</snm><x>,</x><sup>', author)
        author = re.sub(r'([ >,\(])([A-Za-z0-9_.-]+@)([^\s<>]+?)(?!</a>)(,|;|\.? |$|<|\))',
                        r'\1<a href="mailto:\2\3" title="mailto:\2\3">\2\3</a>\4', author, flags=re.IGNORECASE)
        author = re.sub(r'(<a href="mailto([^>]+)>((?:[^<]*|<(?!/a))+)</a></sup>)', r'</sup><sup>\1', author,
                        flags=re.IGNORECASE)
        author = re.sub(r'(\s\&bull\;)', r'</au><x>\1</x><au>', author)
        author = re.sub(r'<deg>(\s*)</deg>', r'\1', author)
        author = re.sub(r'</deg>(\;\s)', r'</deg></au><x>\1</x><au>', author)
        # print(author)
        author = re.sub(r'(\;|\.|\,) <persion>', r'</au><x>\1 </x><au><persion>', author)
        # author = re.sub(r'(<x>(;|.|,) and </x>)', r'\1<au>', author)
        author = re.sub('<persion>', '', author, re.DOTALL)
        author = re.sub(r'</persion>', '', author, re.DOTALL)
        author = re.sub(r'<sup>([a-z])([0-9]+)</sup>', r'<sup>\1,\2</sup>', author, flags=re.IGNORECASE)
        author = re.sub(r'(</deg><sup>(?:[^><]+))(\,\s*)(</sup>)', r'\1\3<x>\2</x>', author, flags=re.IGNORECASE)
        superflag = "false"
        if not re.search(r'(<sup>|mailto|orcid\.org)', author, re.IGNORECASE | re.DOTALL):
            author = re.sub(rf'{self.AffSymbolDesigs}', r'<sup>\g<0></sup>', author, flags=re.IGNORECASE | re.DOTALL)
            author = re.sub(r'([0-9]+)', r'<sup>\1</sup>', author, flags=re.IGNORECASE | re.DOTALL)
            author = re.sub(r'(&#)<sup>([0-9]+)</sup>;', r'\1\2;', author, flags=re.IGNORECASE | re.DOTALL)
            superflag = 'true'
        author = re.sub(r'(<sup>)((?:[^<]+|<(?!/sup>))+)(</sup>)', self.replace_sup_tags, author,
                        flags=re.IGNORECASE | re.DOTALL)
        # author = re.sub(r'(<sup[^>]*>)(\[?)(.*?)(\]?</sup>)', r'\1' + self.crosslinkauthor(r'\3') + r'\4', author, flags=re.IGNORECASE)
        # author = re.sub(r'((?:\*)+)', self.crosslinkauthor(r'\1'), author) if not re.search(r'(lowast">|"\*+">|\#\*+)', author, re.IGNORECASE) else author
        # author = re.sub(r'((?:\*)+)(\,)', self.crosslinkauthor(r'\1') + r'\2', author)
        author = re.sub(r'(' + self.role + ')', r'<role>\1</role>', author, flags=re.IGNORECASE)
        author = re.sub(r'(\(' + self.pronouns + r'\))', r'<pronouns>\1</pronouns>', author, flags=re.I | re.S)
        author = re.sub(r'([ >,])(' + self.gen + r')(?=[^a-zA-Z])', r'\1<gen>\2</gen>', author, flags=re.S)
        author = re.sub(r'(\s(\&|\&amp;)\s)', r'</au><x>\1</x><au>', author)
        author = re.sub(r'<x>((\;|\,|.|)\sand\s)<\/x>', r'</au><x>\1</x><au>', author)
        author = re.sub(r'(\s?\&middot\;\s?)', r'</au><x>\1</x><au>', author)
        author = re.sub(r'(\&[a-z]+)\;', r'\1SEMICOLON ', author)
        author = re.sub(r'(\son\sbehalf\sof\sthe\s|by:\s*|authors?\s*(?:names)?:\s*)', r'<x>\1</x>', author, flags=re.IGNORECASE)
        author = re.sub(r'<sup>(\s+)</sup>', ' ', author, flags=re.IGNORECASE)
        author = re.sub(r'</?sub>', '', author, flags=re.IGNORECASE)
        author = re.sub(r'(,\s?)</p>', r'<x>\1</x></p>', author, flags=re.IGNORECASE)
        author = re.sub(r',</deg>', '</deg>,', author, flags=re.IGNORECASE)
        author = re.sub(r'</deg>([\s,]+)(\((?:Hons|ASCP)\)|$degree)([\s,]+)', r'\1\2</deg>\3', author, flags=re.IGNORECASE)
        hyper_text_val = {}
        hyper_text_count = 0
        au_val = author
        for match in re.finditer(r'<a\s+href="\#?((cor|aff)[^"]+)"', au_val, re.IGNORECASE):
            hyper_text_val[match.group(1)] = 1
        hyper_text_count = len(hyper_text_val)
        author = re.sub(r'(</a>(?:<x>(?:[^><]+)</x>)?</sup>(?:\,?\s*<a href[^><]+>[^><]+</a>)?)(\s?[,;]\s*)([A-Za-z]+)',
                        r'\1</au><x>\2</x><au>\3', author) if hyper_text_count > 1 else author
        # author = re.sub(
        #     r'<au>(?:(?:[^<]*|<(?!/(au|p)))+)</p>',
        #     lambda match: re.sub(
        #         r'((?:<au>(?:(?:[^<]*|<(?!/(a|sup|p)))+)))',
        #         lambda aus_match: re.sub(
        #             r'(\s?[,;]\s|(?:and\s))([A-Za-z]+)',
        #             r'</au><x>\1</x><au>\2',
        #             aus_match.group(0),
        #             flags=re.IGNORECASE | re.DOTALL
        #         ),
        #         match.group(0),
        #         flags=re.IGNORECASE | re.DOTALL
        #     ),
        #     author,
        #     flags=re.IGNORECASE | re.DOTALL
        # )
        author = self.process_author_string(author)
        author = re.sub(r"<\/sup> ([A-Z])", r"</sup></au> <au>\1", author)

        author = re.sub(
            r'(</a>(?:<x>(?:[^><]+)</x>)?</sup>(?:(?:<x>(?:[^><]+)</x>)?\,?\s*<a href[^><]+>[^><]+</a>)?)(\s?[,;]\s*)([A-Za-z]*)',
            r'\1</au><x>\2</x><au>\3', author)
        # while re.search(r'(([^,<>]+)([\,]\s+)([^,<>]+\s*))', author, re.IGNORECASE):
        #     au_match_content = re.search(r'(([^,<>]+)([\,]\s+)([^,<>]+\s*))', author, re.IGNORECASE).group(0)
        #     au_match_content = re.split(' ', au_match_content)
        #     author = re.sub(r'(([^,<>]+)([\,]\s+)([^,<>]+\s*))', r'\2</au><x>\3</x><au>\4', author) if len(
        #         au_match_content) > 2 else author
        # pattern = re.compile(r'(([^,<>]+)([\,]\s+)([^,<>\s]+\s*)(?![^<]*<\/(deg|sup)>))', re.IGNORECASE)
        # while pattern.search(author):
        #     au_match_content = pattern.search(author).group(0)
        #     au_match_content = re.split(' ', au_match_content)
        #     author = pattern.sub(r'\2</au><x>\3</x><au>\4', author) if len(au_match_content) > 2 else author
        pattern = re.compile(r'(([^,<>]+)([\,]\s+)([^,<>\s]+\s*)(?![^<]*<\/(deg|sup)>))', re.IGNORECASE)
        while pattern.search(author):
            au_match_content = pattern.search(author).group(0)
            au_match_content = re.split(' ', au_match_content)
            if len(au_match_content) > 2:
                author = pattern.sub(r'\2</au><x>\3</x><au>\4', author)
            else:
                break  # Exit the loop if no valid match is found or the condition isn't met
        if not re.search(r'</au>', author, re.IGNORECASE | re.DOTALL):
            author = re.sub(r'([,|;]? [ua]nd |[,|;| ] )([^, ]+) ((?:[^,. |^;. ]|&[^&;];)+)', r'</au><x>\1</x><au>\2 \3',
                            author, flags=re.IGNORECASE)
        author = re.sub(r'(</sup>)(\s*and\s+)([A-Za-z]+)', r'\1</au><x>\2</x><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'([A-Za-z]+)(\,\s+and\s+)([A-Za-z]+)', r'\1</au><x>\2</x><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(
            r'(<au><fnm>(?:[^><]+)</fnm>\s<ausurname>(?:[^><]+)</ausurname>)(, )(<fnm>(?:[^><]+)</fnm>\s<ausurname>(?:[^><]+)</ausurname>)',
            r'\1</au><x>\2</x><au>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'((?<!<au>)(?:[A-Za-z]+)\.?\s(?:[A-Za-z]+)(?:\.?\s(?:[A-Za-z]+))?)(\;\s*)',
                        r'\1</au><x>\2</x><au>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<au>(?:[A-Za-z]+)\.?\s(?:[A-Za-z]+)(?:\.?\s(?:[A-Za-z]+))?)(\,\s*)', r'\1</au><x>\2</x><au>',
                        author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<deg>(?:[^<]*|<(?!\/deg))+<\/deg>)(\,\s)([A-z][a-z]*)', r'\1</au><x>\2</x><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'</au>(<x>(?:[^<]*|<(?!/x))+</x>)<au>(<deg>(?:[^<]*|<(?!/deg))+</deg>)', r'\1\2', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'<x></au>(<x>([^><]+)</x>)<au></x>', r'\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'([ ,]+)(</au><x>)', r'\2\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</x>)(<au>)(and\s)', r'\3\1\2', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<au>)([;,\s]+(?:and\s)?)', r'<x>\2</x>\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'\s*[\(\)]\s*', r'<x>\g<0></x>', author, flags=re.IGNORECASE | re.DOTALL)
        author = self.author_pronouns(author)
        if not re.search(r'</au>', author, re.IGNORECASE | re.DOTALL):
            author = re.sub(r'(</sup>(?:<a href[^><]+>[^><]+</a>))(and\s|\;\s|\,\s)', r'\1</au><x>\2</x><au>', author,
                            flags=re.IGNORECASE)
        author = re.sub(
            r'(</deg><sup>(?:[^><]+|<a href[^><]+>[^><]+</a>)</sup>(?:<x>(?:[^><]+)</x>)?(?:<a href[^><]+>[^><]+</a>))(and\s|\;\s|\,\s)([A-Za-z]+)',
            r'\1</au><x>\2</x><au>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'<x> </x>(href=|title=)', r' \1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</sup>)(\,\s)([A-Z]+)', r'\1</au><x>\2</x><au>\3', author, flags=re.DOTALL)
        # author = re.sub(r'((?:[^<]*|<(?!/x))+)(\sand\s)([A-Za-z]+)', r'\1</au><x>\2</x><au>\3', author,
        #                 flags=re.I | re.S)
        # author = re.sub(r'((?:[^<]*|<(?!/x))+)(\sand\s)([A-Za-z]+)', r'\1</au><x>\2</x><au>\3', author,
        #                 flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<fnm>(.*?)</fnm>)', lambda m: re.sub(r'</?deg>', '', m.group(1)), author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</xref_fn>)([, ]+)(<fnm>)', r'\1</au><x>\2</x><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</sup>)(\sand\s)(<fnm>)', r'\1</au><x>\2</x><au>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</sup>|</a>)(\s)([A-Za-z]+)', r'\1</au><x>\2</x><au>\3', author)
        author = re.sub(r'</au><x>((?:[^<]*|<(?!/x))+)</x><au>([^>]+)</deg>', r'\1\2</deg>', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<sup>|,)+(\s*)($AffSymbolDesigs+)(</sup>)?(\s*)([A-z|\&])', r'\1\2\3\4</au><x>\5</x><au>\6',
                        author)
        author = re.sub(r'(<sup>[A-z0-9])+(\s*)(</sup>)?(\s*)([A-z|\&])', r'\1\2\3</au><x>\4</x><au>\5', author)
        author = re.sub(r'</au></au>', r'</au>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'<au><au>', r'<au>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<a href="\#[^\"]+" title="[^\"]+">)\1([^><]+)</a>(</a>)', r'\1\2\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</sup>)(\,?\;?\s*)(<collab>(?:(?:[^<]*|<(?!/collab))+)</collab>)(\;?\,?\s*)([A-Z])',
                        r'\1</au><x>\2</x><au>\3</au><x>\4</x><au>\5', author, flags=re.IGNORECASE | re.DOTALL)
        consortiumflag = "false"
        if re.search(r'consortiumauthors', author, re.IGNORECASE | re.DOTALL):
            consortiumflag = "true"
        author = re.sub(r'<p\s+[^<>]*class="?(?:authors|consortiumauthors)[^<>]*>(.*?)</p>', r'\1', author,
                        flags=re.DOTALL)
        author = re.sub(r'(\,\s)(<deg>)', r'<x>\1</x>\2', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'</au><au>(<gen>|<deg>)', r'\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'\&amp<x>([;\s]+)</x>', r'<x>\&amp\1</x>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(
            r'(<au>)((?:[A-Z\-\&;]+)\s(?:[A-Z\-\&;]+))(<x>(?:[^><]+)</x>)((?:[A-Z\-\&;]+)\s(?:De\s)?(?:[A-Z\-\&;]+))(<x>(?:[^><]+)</x>)(</au>)',
            r'\1\2</au>\3<au>\4\6\5', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'</au><au>(<a href="mailto(?:[^><]+)>)', r'\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'<x><x>([^><]+)</x></x>', r'<x>\1</x>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</au>)(<x>([^><]+)</x>)', r'\2\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'SEMICOLON', r'\;', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</snm>)(<x>(?:[^><]+)</x>)(<collab>)', r'\1\2</au><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</a></sup>)(\s*)(<fnm>)', r'\1<x>\2</x></au><au>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</au>)(\s)(<au>)', r'\1<x>\2</x>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<a(?:[^><]+)>(?:[^><]+)</a>)(\s)(\w+\.?\s\w+\s\w+)', r'\1<x>\2</x></au><au>\3', author,
                        flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</a>\s*</sup>)(\s([^\s]+)\s([^\s]+)\s([^\s]+)<x>[^><]+</x>\s*<deg>)', r'\1</au><au>\2',
                        author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</x>(?:</au>)?)(<au>)([ ,]+)', r'\3\1\2', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(</sup>|</a>)(\s*\()(<a[^><]+>)', r'\1<x>\2</x>\3', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(\)\s*)(<x>)', r'\2\1', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(\)\s*)$', r'<x>\1</x>', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'\,\,(?!\s)', r',', author, flags=re.IGNORECASE | re.DOTALL)
        author = re.sub(r'(<collab>)((?:[^<]+|<(?!/collab>))+)(</collab>)',
                        lambda m: re.sub(r'</?(deg|x|au)>', '', m.group(0)), author, flags=re.IGNORECASE | re.DOTALL)
        if superflag == 'true':
            author = re.sub(r'(<sup>|</sup>)', '', author, flags=re.IGNORECASE | re.DOTALL)
        authors = author
        authors = re.sub(r'\n', '', authors)
        authors = '<au>' + authors + '</au>'
        authors = re.sub(r'<(\/?au)><\1>', r'<\1>', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(\,? and\s?)(</au>)', r'\2<x>\1</x>', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(\,? and\s?)(<x>)', r'\2\1', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'</au><au>(<deg>[^><]+</deg>\s*<role>[^><]+</role></au>)', r'\1', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'<x> \(</x>(http:\/\/orcid\.org(?:.*?))<x>\) </x>', r'<x> </x>(\1)<x> </x>', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(\(http:\/\/orcid\.org/[^)]+\))(\sand\s|\;\s|\,\s)(<fnm>)', r'\1</au><x>\2</x><au>\3',
                         authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(\(http:\/\/orcid\.org/[^)]+\))(<x>[^<]+</x>)(<fnm>)', r'\1\2</au><au>\3', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'(<\/ausurname>\,?<xref_fn>(.*?)<\/xref_fn>)(?:<x> \(<\/x>)(http:\/\/orcid\.org/[^<]+)(?:<x>\)<\/x><x> and <\/x>)(<\/au><au>)(<fnm>)',
            r'\1<x> </x>(\3)<x> and </x>\4\5', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'(<\/ausurname>)(?:<x> \(<\/x>)(http:\/\/orcid\.org/[^<]+)(?:<x>\)<\/x><x> and <\/x>)(<\/au><au>)(<fnm>)',
            r'\1<x> </x>(\2)<x> and </x>\3\4', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'(<\/ausurname><x>[^><]+<\/x>)(<\/au><au>)((?:<x>\(<\/x>))(http:\/\/orcid\.org/[^<]+)((?:<x>\) <\/x>and ))(<fnm>)',
            r'\1(\3)<x> and </x>\2\4\5', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'(<\/ausurname><x>[^><]+<\/x>)(<\/au><au>)(?:<x>\(<\/x>)(http:\/\/orcid\.org/[^<]+)(?:<x>\) <\/x>)(<fnm>)',
            r'\1(\3)<x> </x>\2\4', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(<a href="mailto[^><]+>[^><]+</a><x>[^><]+</x>)(([A-Z][a-z]+)\s([A-Za-z]+))',
                         r'\1</au><au>\2', authors, flags=re.DOTALL)
        authors = re.sub(r'(<x>[^<]+</x>)</au><au>(<a href="\#[^>]+" title="[^>]+">[^<]+</a>)(<x>[^<]+</x>)</au>',
                         r'\2\1\3</au>', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'</x><x>', '', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'<au><a href="\#." title=".">.</a>(<x>[^<]+</x>)</au>', r'\1', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(<a href="ORCHID[^"]+" [^>]+>(?:[^>]+)</a>)<x>([^>]+)</x>([,\s]+)', r'\1<x>\2</x></au><au>',
                         authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'</deg>\.', r'.</deg>', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(\.)<x>', r'<x>\1', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'</role>(\s*)<role>', r'\1', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'(</a>)(\sand\s|\;\s?|\,\s?|\s)(<a href)', r'\1<x>\2</x>\3', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'<x>(\s*\(\s*)</x>(http:\/\/orcid\.org\/\s?[0-9A-Z\-]+)<x>(\s*\)\s*)</x>', r'\1\2\3', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(r'</au><au>(<deg>(?:[^<]*|<(?!/deg))+</deg>\s*<gen>[^><]+</gen></au>)', r'\1', authors,
                         flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'</au><au>(<deg(?:[^<]*|<(?!/deg))+</deg>[;,]?\s*(?:<sup>)?(<a[^><]+>[^><]+</a>(?:<sup>)?(<x>[^><]+</x>)?)+(?:</sup>)?(?:<a[^><]+>[^><]+</a>)?(?:</sup>)?(<x>[^><]+</x>)?</au>)',
            r'\1', authors, flags=re.IGNORECASE | re.DOTALL)
        authors = re.sub(
            r'</au><au>((?:<gen>[^><]+</gen>)?(?:<sup>)?(<a[^><]+>[^><]+</a>(?:<sup>)?(<x>[^><]+</x>)?)+(?:</sup>)?(?:<a[^><]+>[^><]+</a>)?(?:</sup>)?(<x>[^><]+</x>)?</au>)',
            r'\1', authors, flags=re.IGNORECASE | re.DOTALL)
        # while re.search(r'(<au>.*?</au>)', authors):
        #     match = re.search(r'(<au>.*?</au>)', authors)
        #     if match:
        #         _ = match.group(1)
        #         authors = re.sub(r'(<au>.*?</au>)', '', authors, 1)
        #
        #         if re.search(r'(<collab-au-start>|<collab-au-end>)', _, re.IGNORECASE | re.DOTALL):
        #             _ = re.sub(r'</?au>', '', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'<x>, </x>', '', _, flags=re.IGNORECASE | re.DOTALL)
        #             tagged_authors += _ + "\n"
        #         else:
        #             _ = re.sub(r'(<au>)(<x>(?:Authors?\s*(?:names)?|by):?\s*</x>)(\,\s)?', r'\2<x>\3</x>\1', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<x>[^><]+</x>)(</au>)', r'\2\1', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'<au><deg>(.*?)</deg>', r'<au><fdeg>\1</fdeg>', _, flags=re.DOTALL)
        #             _ = re.sub(r'((?:<collab>(?:(?:[^<]*|<(?!/collab))+)</collab>)+)(</au>)', r'</au>\1', _,
        #                        flags=re.DOTALL)
        #             init = r'(?:[A-ZÀÁÂÃÅÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝŠŸŽ\;]\.? ?)+'
        #             word = r'(?:(?:[A-ZÀÁÂÃÅÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝŠŸŽ]|&[A-Z][^&; ]+;)[^A-Z0-9<> ]+)'
        #             word1 = r'(?:(?:[A-ZÀÁÂÃÅÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝŠŸŽ]|&[A-Z][^&; ]+;)[^A-Z0-9]+\-)'
        #             init1 = r'(?:[A-ZÀÁÂÃÅÈÉÊËÌÍÎÏÑÒÓÔÕÖØÙÚÛÜÝŠŸŽ]\. ?)+'
        #             cap = r'(?:(?:[A-Z]|&[A-Z][^&; ]+;)[^0-9]+)'
        #
        #             if not re.search(r'<fnm>', _):
        #                 _ = re.sub(rf'({init})<x>([^>]+)</x>', r'\1\2', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(\,? )(<a href=)', r'<x>\1</x>\2', _, flags=re.IGNORECASE | re.DOTALL)
        #
        #             patterns = [
        #                 (rf'({init})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({init} {word})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word})( (?:{init1}){init1})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word} {init})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word}(?: {word})?)( (?:Lee|Chang?))', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word} {init})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word}(?:[\- ]{word})?)( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({init})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word}) ({init})', r'\1<fnm>\2</fnm> '),
        #                 (rf'({word}) ({init}(?:{init})?)', r'\1<fnm>\2</fnm> '),
        #                 (rf'({word1}\-{init})( {word})', r'\1<fnm>\2</fnm>'),
        #                 (rf'({word1}{word}) ({init})', r'\1\2 <fnm>\3</fnm>'),
        #                 (rf'({cap}) ({cap})', r'\1<fnm>\2</fnm> ')
        #             ]
        #             pat_cnt = 0
        #             for pattern, replace in patterns:
        #                 print(pat_cnt)
        #                 print(pattern)
        #                 _ = self.put_fnm(pattern, replace, _)
        #                 pat_cnt += 1
        #             _ = re.sub(r'(<deg>)(</fnm>)', r'\2\1', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(
        #                 r'(<au>)<fnm>([^><]+)</fnm>(\,?\s)([A-Z]\.\s*(?:([A-Z]\.))?)(\s*\,?(?:(?:<a href[^><]+>[^><]+</a>)?<sup>))',
        #                 r'\1\2\3<fnm>\4</fnm>\6', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'</fnm> ((?:[A-Z]\.? ?)+) ([A-Za-z]+)?', r' \1</fnm> \2', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'<fnm>([^\s]+)(\s[^\s]+\s[A-Z\.]+)</fnm>(\s<sup>)', r'<fnm>\1</fnm>\2\3', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<au><fnm>[^\s]+)\s([^\s]+)</fnm>(\s*<sup>)', r'\1</fnm><x> </x>\2\3', _)
        #             _ = re.sub(r'(<au><fnm>[^><]+)(</fnm>)(\s[A-Z]\.\-[A-Z])(\s)', r'\1\3\2<x>\4</x>', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<au><fnm>[^><]+</fnm>\s[A-Z]\.(?:[A-Z]\.)*(?:[A-Z])*)</au><x>\.', r'\1.</au><x>', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<au>)<fnm>([^><]+)</fnm>(\s)([A-Z]\.(?:[A-Z]\.)*)(\s)?', r'\1\2\3<fnm>\4</fnm>\5', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<au>)<fnm>([^><]+)</fnm>(\s)([A-Z]{1,2}\.)(<sup>)', r'\1\2\3<fnm>\4</fnm>\5', _,
        #                        flags=re.DOTALL)
        #             _ = re.sub(
        #                 r'(<role>(?:[^><]+)</role>\s*)([^\s\,]+)(\,?\s[^\s]+)((?:<x>(?:[^><]+)</x>)?\,?\s?<deg>(?:[^><]+)</deg>)',
        #                 r'\1<fnm>\2</fnm>\3\4', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'\,</fnm>\s', r'</fnm><x>, </x>', _, flags=re.IGNORECASE | re.DOTALL)
        #             if not re.search(r'<fnm>', _):
        #                 _ = re.sub(r'f?deg>', r'fnm>', _)
        #             _ = re.sub(r'(, ?)</sup>', r'</sup>\1', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<a\shref=\"\#)([a-z]+)\&([a-z]+)\;(\"\stitle=\")([a-z]+)\&([a-z]+)\;(\">)',
        #                        r'\1\2\3\4\5\6\7', _, flags=re.IGNORECASE)
        #             while re.search(r'<x>(<x>[^>]*</x>)</x>', _):
        #                 _ = re.sub(r'<x>(<x>[^>]*</x>)</x>', r'\1', _)
        #             _ = re.sub(r'(<x>[^>]*</x>)(</au>)', r'\2\1', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(</ausurname>)([, ]*)(<gen>(.*?)</gen>)(<x>[^>]*</x>)(<fnm>)',
        #                        r'\1<x>\2</x>\3</au>\5<au>\6', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'([, ]+)(<(?:fnm|role|gen|deg|sup)>)', r'<x>\1</x>\2', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(</(?:fnm|role|gen|deg)>)([, ]+)', r'\1<x>\2</x>', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(<au>)(<x>[^><]+</x>)', r'\2\1', _, flags=re.IGNORECASE | re.DOTALL)
        #             _ = re.sub(r'(\()((?:&\#(?:[0-9]+)\;)+)(\))', r'<x>\1</x>\2<x>\3</x>', _,
        #                        flags=re.IGNORECASE | re.DOTALL)
        #             tagged_authors += _ + "\n"
        tagged_authors = authors
        tagged_authors = re.sub(
            r'<au>(and )?(on behalf of (?:the )?)(<collab>)((?:[^<]+)</collab>(?:(?:<sup>)?<a\shref="[^>]+>.*?</a>(?:</sup>))?)</au>',
            r'<x>\1</x>\3\2\4', tagged_authors, flags=re.I | re.S)
        tagged_authors = re.sub(r'<au>(<collab>(?:[^<]+)</collab>(?:(?:<sup>)?<a\shref="[^>]+>.*?</a>(?:</sup>))?)</au>',
                               r'\1', tagged_authors, flags=re.I | re.S)
        tagged_authors = re.sub(r' +(</x>)?$', r'\1', tagged_authors)
        tagged_authors = re.sub(r'(\,? and\s?)(</au>)', r'\2<x>\1</x>', tagged_authors, flags=re.I | re.S)
        tagged_authors = re.sub(r'<x></x>', '', tagged_authors)
        tagged_authors = re.sub(r'<span class="snm">(\s+)</span>', '', tagged_authors, flags=re.I | re.S)
        tagged_authors = re.sub(r'<span class="snm">(\;+)</span>', r'\1', tagged_authors, flags=re.I | re.S)
        tagged_authors = re.sub(r'(<span class="snm">[^\,]+)(\,\s*)(</span>)', r'\1\3\2', tagged_authors,
                               flags=re.I | re.S)
        tagged_authors = re.sub(r'\s*$', '', tagged_authors, flags=re.I | re.S)
        while re.search(r'<x>(<x>[^>]*</x>)</x>', tagged_authors):
            tagged_authors = re.sub(r'<x>(<x>[^>]*</x>)</x>', r'\1', tagged_authors)
        if comment:
            tagged_authors += comment
            comment = None
        # if consortiumflag == 'true':
        #     gauthor = f'<p class="consortiumauthors">{tagged_authors}</p>'
        # else:
        #     gauthor = f'<p class="authors">{tagged_authors}</p>'
        gauthor = tagged_authors
        # gauthor = re.sub(r'\s+</p>', '</p>', gauthor, flags=re.I | re.S)
        gauthor, only_authors = self.author_post_process(gauthor)
        au_count = 0
        pattern = r"<au>((?:[^<]*|<(?!/au))+)</au>"
        while re.search(pattern, gauthor, re.I):
            au_count = au_count + 1
            gauthor = re.sub(pattern, rf'\1<span style="color: blue;">[AU{au_count}]</span>', gauthor, 1, re.IGNORECASE)
        aut_dict = {}
        au_count = 0
        pattern = re.compile(r"<au>(?P<name>((?:[^<]*|<(?!/au))+))</au>")
        matches = pattern.findall(only_authors)
        for match in matches:
            name, full_name = match
            au_count += 1
            aut_dict[au_count] = name
        return gauthor, aut_dict

    def crosslinkauthor(self, link):
        if not re.search(r'(mailto|th)', link):
            if re.search(r'[\s.,]', link):
                link = re.sub(r'([\s.,]+)', r'</sup><x>\1</x><sup>', link)
                link = re.sub(rf'((?:[0-9]+)|(?:[a-z]))(\${self.AffSymbolDesigs})', r'\1</sup><sup>\2</sup>', link)
                link = re.sub(rf'(\${self.AffSymbolDesigs})((?:[0-9]+)|(?:[a-z]))', r'\1</sup><sup>\2</sup>', link,
                              flags=re.IGNORECASE)
            elif len(link) != 1:
                if re.search(rf'(\${self.AffSymbolDesigs})', link, flags=re.IGNORECASE):
                    if not re.search(r'^(\&[^;]+;)', link):
                        link = re.sub(r'([a-z])', r'<sup>\1</sup>', link)
                    if not re.search(r'^(\&[^;]+;)', link):
                        link = re.sub(r'([0-9])', r'<sup>\1</sup>', link)
                    link = re.sub(rf'(\${self.AffSymbolDesigs})', r'<sup>\1</sup>', link, flags=re.IGNORECASE)
                else:
                    link = re.sub(r'([0-9a-z]+\)?)', r'<sup>\1</sup>', link)

            link = re.sub(r'</sup><x>([\s.,]+)</x><sup>$', r'</sup><x>\1</x>', link)
            link = re.sub(r'([0-9]+)(<sup>\*</sup>)', r'<sup>\1</sup>\2', link)
            link = "<sup>" + link + "</sup>"
            link = re.sub(rf'(<sup>)([0-9a-z]+\)?|\${self.AffSymbolDesigs}\)?|(\&[a-z]+\;)|\&+|.)(</sup>)',
                          lambda m: f'<a href="#{self.remove_text(m.group(2), "")}">{m.group(2)}{m.group(3)}</a>', link,
                          flags=re.IGNORECASE)
            link = re.sub(r'(<sup>|</sup>|<a href="[^><]+"></a>)', '', link, flags=re.IGNORECASE)

        link = re.sub(r'(<a href="#[0-9a-z]+)\)', r'\1', link, flags=re.IGNORECASE)
        link = re.sub(r'<a href="#[0-9]+"(?:[^>]+)>([0-9]+)</a>', r'\1', link, flags=re.IGNORECASE)
        link = re.sub(r'(<a\shref="#)([^\"]+)">', lambda m: f'{m.group(1)}{m.group(2)}" title="{m.group(2)}">', link,
                      flags=re.IGNORECASE)
        link = re.sub(
            r'<a\shref="#aff([0-9]+)"\stitle="aff[0-9]+">(?:(?:[^<]+|<(?!/a>))+)</a>(\&ndash\;|[\–\—\-])<a\shref="#aff([0-9]+)"\stitle="aff[0-9]+">(?:(?:[^<]+|<(?!/a>))+)</a>',
            lambda
                m: f'<a href="#aff{m.group(1)}" title="{" ".join([f"aff{i}" for i in range(int(m.group(1)), int(m.group(3)) + 1)])}">{m.group(1)}{m.group(2)}{m.group(3)}</a>',
            link, flags=re.IGNORECASE)
        return link

    def remove_text(self, fullstring, remove_string):
        return re.sub(re.escape(remove_string), '', fullstring)

    def put_fnm(self, pattern, replace, text):
        if not re.search(r'<fnm>', text):
            text = re.sub(rf'(<au>\s*(?:<(?:a|fdeg|role|sup)[^><]*>.*?</(?:a|fdeg|role|sup)>)?){pattern}',
                          lambda m: eval(replace), text)
            text = re.sub(r'(,)(</fnm>)', r'\2\1', text)
        return text

    def replace_sup_tags(self, match):
        open_tag = match.group(1)
        content = match.group(2)
        close_tag = match.group(3)
        full_content = match.group(0)
        if re.search(r'([*†$‡§¶@#^+])(\1)*', content, re.IGNORECASE | re.DOTALL):
            content = re.sub(r'([*†$‡§¶@#^+])(\1)*', r',\g<0>', content, flags=re.IGNORECASE | re.DOTALL)
            content = re.sub(r'^(,)', '', content, flags=re.IGNORECASE | re.DOTALL)
            content = re.sub(r'[,]+', ',', content, flags=re.IGNORECASE | re.DOTALL)
            content = re.sub(r'([*†$‡§¶@#^+])([a-z0-9])', r'\1,\2', content, flags=re.IGNORECASE | re.DOTALL)
        full_content = open_tag + content + close_tag
        return full_content

    def author_pronouns(self, author):
        def replace_pronouns(match):
            pronouns_cont = match.group(0)
            pronouns_cont = re.sub(r'(<x>|</x>|<au>|</au>)', '', pronouns_cont, flags=re.I | re.S)
            return pronouns_cont

        def replace_deg(match):
            deg_cont = match.group(0)
            if not re.search(r'<a href', deg_cont, flags=re.I | re.S):
                deg_cont = re.sub(r'<x>(.*?)</x>', r'\1', deg_cont, flags=re.I | re.S)
            if re.search(r'</au>', deg_cont, flags=re.I | re.S):
                deg_cont = re.sub(r'</au>', '', deg_cont, flags=re.I | re.S)
                deg_cont += '</au>'
            if re.search(r'<au>', deg_cont, flags=re.I | re.S):
                deg_cont = re.sub(r'<au>', '', deg_cont, flags=re.I | re.S)
                deg_cont += '<au>'
            return deg_cont
        # author = re.sub(r'<pronouns>(((?!</?pronouns>).)*)</pronouns>', r"\1", author, flags=re.I | re.S)
        # author = re.sub(r'<deg>(((?!</?deg>).)*)</deg>', r"\1", author, flags=re.I | re.S)
        # author = re.sub(r'<x>(((?!</?x>).)*)</x>', r"\1", author, flags=re.I | re.S)
        # author = re.sub(r'<role>(((?!</?role>).)*)</role>', r"\1", author, flags=re.I | re.S)
        return author

    def process_author_string(self, author):
        def process_au_content(au_content):
            # Replace occurrences of ", " or "; " or " and " followed by a name with a closing and opening tag
            return re.sub(r'(\s?[,;]\s|\s?and\s)([A-Za-z]+)',
                          r'</au><x>\1</x><au>\2',
                          au_content,
                          flags=re.IGNORECASE | re.DOTALL)

        def process_match(match):
            au_content = match.group(0)
            return re.sub(r'(<au>[^<]*)', lambda m: process_au_content(m.group(0)), au_content,
                          flags=re.IGNORECASE | re.DOTALL)

        return re.sub(r'(<au>.*?</p>)', process_match, author, flags=re.IGNORECASE | re.DOTALL)

    def predict_degrees(self, text, degrees):
        text = re.sub(rf'(\,|\s)({self.degree})(?=([^a-zA-Z]|$))', r'\1<deg>\2</deg>\3', text)
        text = re.sub(r'<([^>]+)>(((?!</?\1>).)*)</\1>', "", text, flags=re.I | re.S | re.DOTALL)
        text = re.sub(r'\s+,', ',', text)
        text = re.sub(" and ", "", text)
        text = re.sub(r'\s+\.', '.', text)
        text = re.sub(r'\s+;', ';', text)
        text = re.sub(r'\s+$', '', text)
        text = re.sub(r',(\s|),', ',', text)
        text = re.sub(r';(\s|);', ';', text)
        text = re.sub(r'\.(\s|)\.', '.', text)
        text = re.sub(r',(\s|)\.', ',', text)
        text = re.sub(r',(\s|);', ',', text)
        text = re.sub(r'\.(\s|),', ',', text)
        text = re.sub(r'\.(\s|);', ';', text)
        text = re.sub(r';(\s|)\.', ';', text)
        text = re.sub(r';(\s|),', ';', text)
        # text = re.sub('^', '1; ', text)
        # text = re.sub('$', ', 0', text)
        use_gpu = False
        warnings.filterwarnings("ignore", message="User provided device_type of 'cuda', but CUDA is not available. Disabling")
        try:
            use_gpu = spacy.require_gpu()
            if use_gpu:
                print("GPU is available, using GPU.")
            else:
                print("GPU not available, using CPU.")
        except ValueError as e:
            print("Cannot use GPU. Falling back to CPU.")
            use_gpu = False
        # nlp = spacy.load("D:\\mProjects\\BreakDown\\venv\\Lib\\site-packages\\en_core_web_trf\\en_core_web_trf-3.7.3")
        # nlp = spacy.load("en_core_web_trf")
        # nlp = spacy.load("en_core_web_md")
        nlp = spacy.load("D:\\mProjects\\BreakDown\\venv\\Lib\\site-packages\\en_core_web_lg\\en_core_web_lg-3.7.1")
        doc = nlp(text)
        other_entities = set()
        person_entities = set()
        degrees_list = degrees.split("|")
        for ent in doc.ents:
            ent_text = ent.text
            ent_label = ent.label_
            ent_list = ent_text.split()
            if ent.text in degrees_list or ent.text in other_entities:
                continue
            if ent.label_ in ["PERSON", "NORP"]:
               person_entities.add(ent.text)
            if ent.label_ not in ["PERSON", "NORP"] and len(ent_list) > 1:
                result = self.check_all_sentence_case(ent_list)
                if result is True:
                    person_entities.add(ent.text)
                    continue
            if ent.label_ not in ["PERSON", "NORP"]:
                if ent.label_ == "GPE" and ent.text.isupper():
                    other_entities.add(ent.text)
                elif ent.label_ in ["TITLE", "WORK_OF_ART"]:
                    other_entities.add(ent.text)
                elif ent.label_ in ["ORG"] and ent.text.isupper():
                    other_entities.add(ent.text)
                else:
                    person_entities.add(ent.text)
        degree_list = list(other_entities)  # Convert set to list
        person_list = list(person_entities)  # Convert set to list
        return degree_list, person_list

    def is_sentence_case(self, s):
        return s == s.capitalize() and s[1:].islower()

    def check_all_sentence_case(self, arr):
        return all(self.is_sentence_case(item) for item in arr)

    def author_post_process(self, gauthor):
        #</deg><span class="snm">
        gauthor = re.sub('</au><x>; </x><au><span class="fnm">Bsc; ', '<x>; </x><deg>Bsc</deg></au><x>; </x><au><span class="fnm">', gauthor, flags=re.I | re.S)
        gauthor = re.sub('</deg><span class="snm">', '</deg></au><x> </x><au><span class="snm">', gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'(</sup>)(\s+)(<a\s+href)', r'\1<x>\2</x>\3', gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'>\;<', r'><x>;</x><', gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'(<x>\;</x>)(<a href[^>]+>[^>]+</a>)', r'\2\1', gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'(\,\s*)(</au>)', r'<x>\1</x>\2', gauthor)
        gauthor = re.sub(r'(<x>[^>]*</x>)(</au>)', r'\2\1', gauthor)
        gauthor = re.sub(r'(\s+)(</span>)', r'\2<x>\1</x>', gauthor, flags=re.I | re.S)
        gauthor = re.sub(
            r'</au>(<x>[^>]*</x>)<au>(<role>[^<]*</role>)(?!(?:<x>[^>]*</x>)?(<role>[^<]*</role>)?(?:<x>[^>]*</x>)?<span class="(s|f)nm)',
            r'\1\2', gauthor)
        gauthor = re.sub(r'</x><x>', '', gauthor)
        if re.search(r'(<span class\=\"role\">author<\/span>)(<\/au>)(<\/p>)', gauthor, flags=re.I | re.S):
            gauthor = re.sub(r'(<span class\=\"role\">author<\/span>)(<\/au>)(<\/p>)', r'\2\1\3', gauthor,
                             flags=re.I | re.S)
        gauthor = re.sub(r'</au>[,\s]*<collab-au-end>collab-au-end</collab-au-end>',
                         r'</au><collab-au-end>collab-au-end</collab-au-end>', gauthor, flags=re.I | re.S)
        only_authors = gauthor
        gauthor = re.sub(r'<pronouns>(((?!</?pronouns>).)*)</pronouns>', r"\1", gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'<deg>(((?!</?deg>).)*)</deg>', r"\1", gauthor, flags=re.I | re.S)
        gauthor = re.sub(r'<role>(((?!</?role>).)*)</role>', r"\1", gauthor, flags=re.I | re.S)
        gauthor = re.sub("<x>", "", gauthor, re.I| re.S)
        gauthor = re.sub("</x>", "", gauthor, re.I| re.S)
        gauthor = re.sub(r"\.</au>$", "</au>.", gauthor, re.I)
        gauthor = re.sub("</span><span class=\"snm\">", " ", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub(r"<span class=\"fnm\">", "", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub(r"<span class=\"snm\">", "", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub("</span>", "", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub(r"(<a>|<a ([^>]+)>)", "", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub("</a>", "", gauthor, re.I | re.S | re.DOTALL)
        gauthor = re.sub("<x>", "", gauthor, re.I| re.S | re.DOTALL)
        gauthor = re.sub("</x>", "", gauthor, re.I| re.S | re.DOTALL)
        only_authors = re.sub(r'(<a( [^>]+|)>(((?!</?a>).)*)</a>)', "", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<pronouns>(((?!</?pronouns>).)*)</pronouns>', r"\1", only_authors, flags=re.I | re.S)
        only_authors = re.sub(r'<deg>(((?!</?deg>).)*)</deg>(\,|\.|\:|\;| and | \& | |)', "", only_authors, flags=re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<sup>(((?!</?sup>).)*)</sup>', "", only_authors, flags=re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<role>(((?!</?role>).)*)</role>', "", only_authors, flags=re.I | re.S | re.DOTALL)
        only_authors = re.sub("<x>", "", only_authors, re.I| re.S)
        only_authors = re.sub("</x>", "", only_authors, re.I| re.S)
        only_authors = re.sub(r"\.</au>$", "</au>.", only_authors, re.I)
        only_authors = re.sub("</span><span class=\"snm\">", " ", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub("<span class=\"fnm\">", "", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub("<span class=\"snm\">", "", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub("</span>", "", only_authors, re.I | re.S)
        only_authors = re.sub("<x>", "", only_authors, re.I| re.S)
        only_authors = re.sub("</x>", "", only_authors, re.I| re.S)
        only_authors = re.sub("  ", " ", only_authors, re.I | re.S)
        only_authors = re.sub(",,", ",", only_authors, re.I | re.S)
        only_authors = re.sub(", ,", ",", only_authors, re.I | re.S)
        only_authors = re.sub(r"\.\.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r"\. \.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r"\.,", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r"\. ,", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r",\.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r",\.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r"\.;", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r"\. ;", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r";\.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r";\.", ".", only_authors, re.I | re.S)
        only_authors = re.sub(r'<au>(\,|\.|\;| and |and | \&|\&|)(\s+|)', "<au>",
                                only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>and ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\, and ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\. and ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\; and ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\: and ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\, \& ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\. \& ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\; \& ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\: \& ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>\: \& ', "<au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'<au>(\,|\.|\;| and | \&|\&|)(\s+|)</au>', "</au>",
                                only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\, and </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\. and </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\; and </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\: and </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\, \& </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\. \& </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\; \& </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\: \& </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'\: \& </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        only_authors = re.sub(r'(\:|\,|\.|\;|) </au>', "</au>", only_authors, re.I | re.S | re.DOTALL)
        return gauthor, only_authors


# proc_auth = PyLabelAuthor();
# aut, only_authors = proc_auth.author_post_process('<p class="authors"><au><span class="fnm">Wolfgang Emanuel</span><span class="snm">Z\x0000fcrrer</span><sup><a href="#1" title="1">1</a><x>,</x><a href="#2" title="2">2</a></sup><x>, </x><deg>BSc</deg></au><x>; </x><au><span class="fnm">Amelia Elaine</span><span class="snm">Cannon</span><sup><a href="#1" title="1">1</a><x>,</x><a href="#2" title="2">2</a></sup></au><x>; </x><au><span class="fnm">Dariya</span><span class="snm">Ilchenko</span><sup><a href="#1" title="1">1</a><x>,</x><a href="#2" title="2">2</a></sup></au><x>; </x><au><span class="fnm">Bsc; Maria</span><span class="snm">Ines Gaitan</span><sup><a href="#3" title="3">3</a></sup><x>, </x><deg>MD, PhD</deg></au><x>; </x><au><span class="fnm">Tobias</span><span class="snm">Granberg</span><sup><a href="#4" title="4">4</a><x>,</x><a href="#5" title="5">5</a></sup><x>, </x><deg>MD, PhD</deg></au><x>; </x><au><span class="fnm">Fredrik</span><span class="snm">Piehl</span><sup><a href="#5" title="5">5</a><x>,</x><a href="#6" title="6">6</a></sup><x>, </x><deg>MD, PhD</deg></au><x>; </x><au><span class="fnm">Andrew J</span><span class="snm">Solomon</span><sup><a href="#7" title="7">7</a></sup><a href="#*" title="*">*</a><x>, </x><deg>MD</deg></au><x>; </x><au><span class="fnm">Benjamin Victor</span><span class="snm">Ineichen</span><x>, </x><deg>MD, PhD</deg><sup><a href="#1" title="1">1</a><x>,</x><a href="#2" title="2">2</a></sup><a href="#*" title="*">*</a></au></p>')
# print(only_authors)
