let
    CountryCodes = Text.Split("ARG AUS BRA CAN CHL CHN DEU EGY ESP FRA GBR IDN IND ITA JPN KOR MEX NGA NOR NZL POL SWE TUR USA ZAF", " "),
    IndicatorCodes = {"NY.GDP.PCAP.PP.KD", "NY.GDP.MKTP.KD", "NY.GDP.MKTP.KD.ZG", "FP.CPI.TOTL.ZG",
        "SP.POP.TOTL", "SP.DYN.LE00.IN", "EG.ELC.ACCS.ZS", "IT.NET.USER.ZS", "NE.TRD.GNFS.ZS"},
    MetricMap = Record.FromList({"Income", "GDP", "Growth", "Inflation", "Population", "Life",
        "Electricity", "Internet", "Trade"}, IndicatorCodes),
    Archive = Binary.Buffer(Web.Contents("https://databankfiles.worldbank.org",
        [RelativePath="public/ddpext_download/WDI_CSV.zip", Timeout=#duration(0,0,10,0)])),
    UInt = (bytes as binary) as number =>
        List.Sum(List.Transform(List.Positions(Binary.ToList(bytes)),
            (i) => Binary.ToList(bytes){i} * Number.Power(256, i))),
    ReadMember = (name as text) as binary =>
        let
            End = Binary.Range(Archive, Binary.Length(Archive) - 22, 22),
            CentralStart = if UInt(Binary.Range(End, 0, 4)) = 101010256
                then UInt(Binary.Range(End, 16, 4))
                else error "WDI archive layout changed: expected a ZIP without an archive comment.",
            Entry = (offset as number) as record =>
                let
                    Header = Binary.Range(Archive, offset, List.Min({46, Binary.Length(Archive) - offset})),
                    Signature = UInt(Binary.Range(Header, 0, 4)),
                    NameLength = UInt(Binary.Range(Header, 28, 2)),
                    ExtraLength = UInt(Binary.Range(Header, 30, 2)),
                    CommentLength = UInt(Binary.Range(Header, 32, 2))
                in [
                    Valid = Signature = 33639248,
                    Name = if Signature = 33639248 then Text.FromBinary(Binary.Range(Archive, offset + 46, NameLength), TextEncoding.Utf8) else "",
                    Method = UInt(Binary.Range(Header, 10, 2)),
                    Size = UInt(Binary.Range(Header, 20, 4)),
                    Local = UInt(Binary.Range(Header, 42, 4)),
                    Next = offset + 46 + NameLength + ExtraLength + CommentLength
                ],
            Entries = List.Generate(() => Entry(CentralStart), each [Valid], each Entry([Next])),
            Matches = List.Select(Entries, each [Name] = name),
            Found = if List.Count(Matches) = 1 then Matches{0} else error "Required CSV is missing or duplicated in the WDI archive.",
            Local = Binary.Range(Archive, Found[Local], 30),
            Start = Found[Local] + 30 + UInt(Binary.Range(Local, 26, 2)) + UInt(Binary.Range(Local, 28, 2)),
            Bytes = Binary.Range(Archive, Start, Found[Size]),
            Data = if Found[Method] = 8 then Binary.Decompress(Bytes, Compression.Deflate)
                else if Found[Method] = 0 then Bytes
                else error "Unsupported ZIP compression in WDI archive."
        in Data,
    LicenseData = Table.PromoteHeaders(Csv.Document(ReadMember("WDISeries.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv])),
    SelectedLicenses = Table.SelectRows(LicenseData, each List.Contains(IndicatorCodes, [Series Code])),
    LicensesOK = Table.RowCount(SelectedLicenses) = List.Count(IndicatorCodes)
        and List.AllTrue(List.Transform(SelectedLicenses[License Type], each _ = "CC BY-4.0")),
    Raw = if LicensesOK
        then Table.PromoteHeaders(Csv.Document(ReadMember("WDICSV.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]))
        else error "Selected WDI indicator licences must all be CC BY-4.0. Review source permissions.",
    Selected = Table.SelectRows(Raw, each List.Contains(CountryCodes, [Country Code]) and List.Contains(IndicatorCodes, [Indicator Code])),
    Years = List.Transform({2000..2024}, each Text.From(_)),
    Kept = Table.SelectColumns(Selected, List.Combine({{"Country Code", "Indicator Code"}, Years})),
    PreserveMissing = Table.ReplaceValue(Kept, null, "", Replacer.ReplaceValue, Years),
    Long = Table.Unpivot(PreserveMissing, Years, "Year", "Value"),
    Renamed = Table.RenameColumns(Long, {{"Country Code", "CountryCode"}, {"Indicator Code", "IndicatorCode"}}),
    Metrics = Table.AddColumn(Renamed, "Metric", each Record.Field(MetricMap, [IndicatorCode]), type text),
    Ordered = Table.ReorderColumns(Metrics, {"CountryCode", "Year", "IndicatorCode", "Metric", "Value"}),
    EmptyToNull = Table.ReplaceValue(Ordered, "", null, Replacer.ReplaceValue, {"Value"}),
    Typed = Table.TransformColumnTypes(EmptyToNull, {{"CountryCode", type text}, {"Year", Int64.Type},
        {"IndicatorCode", type text}, {"Metric", type text}, {"Value", type number}}, "en-US"),
    Keys = Table.Group(Typed, {"CountryCode", "Year", "IndicatorCode"}, {{"Count", each Table.RowCount(_), Int64.Type}}),
    ExpectedRows = List.Count(CountryCodes) * List.Count(IndicatorCodes) * List.Count(Years),
    Checked = if Table.RowCount(Table.SelectRows(Keys, each [Count] <> 1)) > 0
        then error "Duplicate country/year/indicator observations."
        else if Table.RowCount(Typed) <> ExpectedRows
        then error "Incomplete country/year/indicator coverage. Missing observations must remain explicit blank rows."
        else Typed,
    ModeColumn = Table.AddColumn(Checked, "SourceMode", each "Live", type text),
    Refreshed = Table.AddColumn(ModeColumn, "LoadedAt", each DateTime.FixedLocalNow(), type datetime)
in Refreshed
