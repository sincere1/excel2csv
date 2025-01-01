const std = @import("std");
const c_xlsx = @cImport({
    @cInclude("xlsxio_read.h");
});

pub fn main() !void {
    var arena = std.heap.ArenaAllocator.init(std.heap.page_allocator);
    defer arena.deinit();

    const allocator = arena.allocator();

    const args = try std.process.argsAlloc(allocator);
    defer std.process.argsFree(allocator, args);

    if (args.len != 3) {
        std.log.warn("Xlsx file path and sheet name are exprected", .{});
        std.process.exit(1);
    }

    const xlsx_file = c_xlsx.xlsxioread_open(args[1]);
    defer c_xlsx.xlsxioread_close(xlsx_file);

    // Write data
    const sheet = c_xlsx.xlsxioread_sheet_open(xlsx_file, args[2], 1);
    var dynamic_string = std.ArrayList(u8).init(allocator);
    defer dynamic_string.deinit();
    var writer = dynamic_string.writer();
    while (c_xlsx.xlsxioread_sheet_next_row(sheet) > 0) {
        while (c_xlsx.xlsxioread_sheet_next_cell(sheet)) |cell| {
            // const temp_value = @as([*:0]const u8, cell);
            try writer.writeAll(std.mem.sliceTo(cell, 0));
            try writer.writeAll("\t");
        }
        try writer.writeAll("\n");
    }

    // Open the file for writing
    const file = try std.fs.cwd().createFile("out_sheet.csv", .{});
    defer file.close();

    // Write the data to the file
    try file.writeAll(dynamic_string.items);
}
