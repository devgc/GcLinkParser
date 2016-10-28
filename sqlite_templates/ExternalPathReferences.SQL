SELECT DISTINCT
link_files.DriveType,
link_files.VolumeLabel,
link_files.DriveSerialNumber,
json_extract(entry_item.value,'$.ExtentionBlocks[0].LongName') AS LongName,
json_extract(entry_item.value,'$.ExtentionBlocks[0].EntryNum') AS EntryNum,
json_extract(entry_item.value,'$.ExtentionBlocks[0].SeqNum') AS SeqNum,
link_files.LocalPath,
json_extract(entry_item.value,'$.ExtentionBlocks[0].CreationTime') AS CreationTime,
json_extract(entry_item.value,'$.ExtentionBlocks[0].AccessTime') AS AccessTime,
link_files.FileSize,
link_files.Source,
json_extract(LnkTrgData,'$.FileEntries') AS LnkTrgData
FROM 
"link_files",
json_each(
    json_extract(LnkTrgData,'$.FileEntries')
) AS entry_item
WHERE
DriveType == 'DRIVE_REMOVABLE;'