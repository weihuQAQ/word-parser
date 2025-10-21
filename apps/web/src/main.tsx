import { createRoot } from "react-dom/client";
import { useRef, useState } from "react";
import { Upload, X } from "lucide-react";
import { BlobReader, TextWriter, ZipReader } from "@zip.js/zip.js";
import { Panel, PanelGroup, PanelResizeHandle } from "react-resizable-panels";

import "@workspace/ui/globals.css";

import { AppSidebar } from "@/components/app-sidebar";
import {
  Breadcrumb,
  BreadcrumbItem,
  BreadcrumbLink,
  BreadcrumbList,
  BreadcrumbPage,
  BreadcrumbSeparator,
} from "@workspace/ui/components/breadcrumb";
import { Button } from "@workspace/ui/components/button";
import { SidebarProvider } from "@workspace/ui/components/sidebar";
import { ScrollArea, ScrollBar } from "@workspace/ui/components/scroll-area";

interface FileEntry {
  filename: string;
  size: number;
  directory: boolean;
}

const App = () => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [fileEntries, setFileEntries] = useState<FileEntry[]>([]);
  const [uploadedFileName, setUploadedFileName] = useState<string>("");
  const [selectedFile, setSelectedFile] = useState<string>("");
  const [openFiles, setOpenFiles] = useState<string[]>([]);
  const [fileContents, setFileContents] = useState<Record<string, string>>({});
  const [zipFile, setZipFile] = useState<File | null>(null);

  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileSelect = async (filepath: string) => {
    setSelectedFile(filepath);
    // 将文件添加到打开列表（如果还没有打开）
    if (!openFiles.includes(filepath)) {
      setOpenFiles([...openFiles, filepath]);
    }

    // 如果文件内容还没有加载，从zip中读取
    if (!fileContents[filepath] && zipFile) {
      try {
        const zipReader = new ZipReader(new BlobReader(zipFile));
        const entries = await zipReader.getEntries();
        const entry = entries.find((e) => e.filename === filepath);

        if (entry && !entry.directory && "getData" in entry) {
          // 尝试读取为文本
          const text = await entry.getData(new TextWriter());
          setFileContents((prev) => ({ ...prev, [filepath]: text }));
        }

        await zipReader.close();
      } catch (error) {
        console.error("Error reading file content:", error);
        setFileContents((prev) => ({
          ...prev,
          [filepath]: "无法读取文件内容",
        }));
      }
    }
  };

  const handleCloseFile = (filepath: string, e?: React.MouseEvent) => {
    e?.stopPropagation();
    const newOpenFiles = openFiles.filter((f) => f !== filepath);
    setOpenFiles(newOpenFiles);

    // 如果关闭的是当前选中的文件，切换到其他文件
    if (filepath === selectedFile) {
      if (newOpenFiles.length > 0) {
        // 切换到前一个或下一个文件
        const currentIndex = openFiles.indexOf(filepath);
        const newIndex = currentIndex > 0 ? currentIndex - 1 : 0;
        setSelectedFile(newOpenFiles[newIndex] || "");
      } else {
        setSelectedFile("");
      }
    }
  };

  const handleFileChange = async (
    event: React.ChangeEvent<HTMLInputElement>
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    console.log("Selected file:", file.name);

    const fileName = file.name;
    const fileExtension = fileName
      .substring(fileName.lastIndexOf("."))
      .toLowerCase();

    if (fileExtension === ".docx" || fileExtension === ".zip") {
      try {
        const zipReader = new ZipReader(new BlobReader(file));
        const entries = await zipReader.getEntries();

        const fileList: FileEntry[] = entries.map((entry) => ({
          filename: entry.filename,
          size: entry.uncompressedSize,
          directory: entry.directory,
        }));

        // 重置所有状态，清除之前的内容
        setUploadedFileName(fileName);
        setFileEntries(fileList);
        setZipFile(file);
        setOpenFiles([]);
        setSelectedFile("");
        setFileContents({});

        await zipReader.close();

        console.log("File processed successfully");
      } catch (error) {
        console.error("Error reading zip file:", error);
      }
    } else {
      console.warn("Unsupported file type:", fileExtension);
    }

    event.target.value = "";
  };

  return (
    <SidebarProvider className="h-full! w-full! min-h-full!">
      <PanelGroup direction="horizontal" className="h-full w-full">
        <Panel defaultSize={20} minSize={15} maxSize={35}>
          <AppSidebar
            fileEntries={fileEntries}
            uploadedFileName={uploadedFileName}
            onFileSelect={handleFileSelect}
          />
        </Panel>
        <PanelResizeHandle className="w-1 bg-border hover:bg-primary/50 active:bg-primary transition-colors cursor-col-resize relative group">
          <div className="absolute inset-y-0 -left-1 -right-1 group-hover:bg-primary/10 transition-colors" />
        </PanelResizeHandle>
        <Panel defaultSize={80} minSize={50}>
          <div className="flex h-full w-full flex-col bg-card">
            <header className="flex h-12 shrink-0 items-center justify-end gap-2 border-b px-4 bg-background">
              <input
                ref={fileInputRef}
                type="file"
                className="hidden"
                onChange={handleFileChange}
                accept=".docx,.zip"
              />
              <Button onClick={handleImportClick} variant="default" size="sm">
                <Upload className="mr-2 h-4 w-4" />
                Import
              </Button>
            </header>

            {/* 文件标签栏 */}
            {openFiles.length > 0 && (
              <div className="border-b bg-secondary/50">
                <ScrollArea className="w-full">
                  <div className="flex items-center h-10">
                    {openFiles.map((filepath) => {
                      const fileName = filepath.split("/").pop() || filepath;
                      const isActive = filepath === selectedFile;
                      return (
                        <button
                          key={filepath}
                          onClick={() => setSelectedFile(filepath)}
                          className={`
                            group relative flex items-center gap-2 px-4 h-full
                            border-r hover:bg-accent/70 transition-colors
                            ${isActive ? "bg-card text-foreground" : "text-muted-foreground"}
                          `}
                        >
                          <span className="text-sm whitespace-nowrap">
                            {fileName}
                          </span>
                          <button
                            onClick={(e) => handleCloseFile(filepath, e)}
                            className="opacity-0 group-hover:opacity-100 hover:bg-muted rounded-sm p-0.5 transition-opacity"
                          >
                            <X className="h-3 w-3" />
                          </button>
                          {isActive && (
                            <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />
                          )}
                        </button>
                      );
                    })}
                  </div>
                  <ScrollBar orientation="horizontal" />
                </ScrollArea>
              </div>
            )}

            {/* 面包屑 */}
            {selectedFile && (
              <div className="flex items-center h-10 border-b px-4 bg-secondary/30">
                <Breadcrumb>
                  <BreadcrumbList>
                    {selectedFile.split("/").map((part, index, arr) => {
                      const isLast = index === arr.length - 1;
                      return (
                        <div key={index} className="flex items-center">
                          <BreadcrumbItem
                            className={index > 0 ? "hidden md:block" : ""}
                          >
                            {isLast ? (
                              <BreadcrumbPage>{part}</BreadcrumbPage>
                            ) : (
                              <BreadcrumbLink href="#">{part}</BreadcrumbLink>
                            )}
                          </BreadcrumbItem>
                          {!isLast && (
                            <BreadcrumbSeparator className="hidden md:block" />
                          )}
                        </div>
                      );
                    })}
                  </BreadcrumbList>
                </Breadcrumb>
              </div>
            )}

            <div className="flex-1 overflow-auto">
              {selectedFile ? (
                <div className="h-full p-6">
                  <div className="bg-accent/40 rounded-lg p-6 h-full">
                    <pre className="text-sm font-mono whitespace-pre-wrap wrap-break-word">
                      {fileContents[selectedFile] || "加载中..."}
                    </pre>
                  </div>
                </div>
              ) : (
                <div className="flex items-center justify-center h-full text-muted-foreground">
                  <div className="text-center">
                    <p className="text-lg">未选择文件11122</p>
                    <p className="text-sm mt-2">请从左侧选择一个文件查看内容</p>
                  </div>
                </div>
              )}
            </div>
          </div>
        </Panel>
      </PanelGroup>
    </SidebarProvider>
  );
};

createRoot(document.getElementById("app")!).render(<App />);
