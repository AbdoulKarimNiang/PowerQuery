
// ReplaceText is the table where the are two columns
- Old is the column with the values to replace
- New is the column with the new values to use as a replacement

List.Accumulate({0..List.Count(ReplaceText[Old])-1}, SelectColumns, (state, current) => Table.ReplaceValue(state, ReplaceText[Old]{current}, ReplaceText[New]{current},Replacer.ReplaceText,{"table_first"}))
