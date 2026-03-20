import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Copy, Check, Table as TableIcon } from "lucide-react";
import { useState } from "react";
import { Button } from "@/components/ui/button";

interface TableViewProps {
    data: any[];
    title?: string;
    maxHeight?: string;
}

const TableView = ({ data, title, maxHeight = "400px" }: TableViewProps) => {
    const [copied, setCopied] = useState(false);

    if (!data || data.length === 0) {
        return (
            <div className="rounded-lg border border-border p-8 text-center text-muted-foreground italic text-sm">
                No data available for table view
            </div>
        );
    }

    // Get all unique keys from the data for headers
    const headers = Array.from(new Set(data.flatMap(row => Object.keys(row))));

    const handleCopyCsv = () => {
        const csvRows = [
            headers.join(','),
            ...data.map(row =>
                headers.map(header => {
                    const val = row[header];
                    const escaped = ('' + (val ?? '')).replace(/"/g, '""');
                    return `"${escaped}"`;
                }).join(',')
            )
        ];
        navigator.clipboard.writeText(csvRows.join('\n'));
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
    };

    return (
        <div className="rounded-lg overflow-hidden border border-border bg-card">
            {title && (
                <div className="flex items-center justify-between px-4 py-2 bg-muted/50 border-b border-border">
                    <div className="flex items-center gap-2">
                        <TableIcon className="w-3.5 h-3.5 text-muted-foreground" />
                        <span className="text-xs font-semibold text-muted-foreground">{title}</span>
                    </div>
                    <Button variant="ghost" size="sm" className="h-7 text-xs" onClick={handleCopyCsv}>
                        {copied ? (
                            <><Check className="w-3 h-3 mr-1 text-green-500" /> Copied CSV</>
                        ) : (
                            <><Copy className="w-3 h-3 mr-1" /> Copy CSV</>
                        )}
                    </Button>
                </div>
            )}
            <div className="overflow-auto scrollbar-thin scrollbar-thumb-muted-foreground/20" style={{ maxHeight }}>
                <Table>
                    <TableHeader className="bg-muted/30 sticky top-0 z-10">
                        <TableRow>
                            {headers.map((header) => (
                                <TableHead key={header} className="text-[10px] font-bold uppercase tracking-wider h-10 px-3 whitespace-nowrap">
                                    {header.replace(/_/g, " ")}
                                </TableHead>
                            ))}
                        </TableRow>
                    </TableHeader>
                    <TableBody>
                        {data.map((row, i) => (
                            <TableRow key={i} className="hover:bg-muted/20 transition-colors border-border/50">
                                {headers.map((header) => (
                                    <TableCell key={header} className="text-[11px] py-2 px-3 whitespace-nowrap text-foreground/80">
                                        {row[header]?.toString() || "-"}
                                    </TableCell>
                                ))}
                            </TableRow>
                        ))}
                    </TableBody>
                </Table>
            </div>
        </div>
    );
};

export default TableView;
