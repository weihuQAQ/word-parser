import { ChevronRight, File, Folder } from "lucide-react";

import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
} from "@workspace/ui/components/collapsible";
import {
  SidebarContent,
  SidebarGroup,
  SidebarGroupContent,
  SidebarGroupLabel,
  SidebarMenu,
  SidebarMenuBadge,
  SidebarMenuButton,
  SidebarMenuItem,
  SidebarMenuSub,
} from "@workspace/ui/components/sidebar";
import { useMemo } from "react";

interface FileEntry {
  filename: string;
  size: number;
  directory: boolean;
}

interface AppSidebarProps {
  fileEntries?: FileEntry[];
  uploadedFileName?: string;
  onFileSelect?: (filepath: string) => void;
}

interface TreeNodeData {
  name: string;
  isDirectory: boolean;
  children: TreeNodeData[];
}

interface TreeNode {
  [key: string]: {
    isDirectory: boolean;
    node: TreeNode;
  } | null;
}

export function AppSidebar({
  fileEntries = [],
  uploadedFileName = "",
  onFileSelect,
}: AppSidebarProps) {
  // 构建树形结构
  const buildTree = (entries: FileEntry[]): TreeNodeData[] => {
    const tree: TreeNode = {};

    entries.forEach((entry) => {
      const parts = entry.filename.split("/").filter(Boolean);
      let current: TreeNode = tree;

      parts.forEach((part, index) => {
        const isLastPart = index === parts.length - 1;

        if (!current[part]) {
          if (isLastPart && !entry.directory) {
            // 这是一个文件
            current[part] = null;
          } else {
            // 这是一个目录
            current[part] = {
              isDirectory: true,
              node: {},
            };
          }
        }

        if (current[part] !== null && !isLastPart) {
          current = current[part]!.node;
        }
      });
    });

    const convertToArray = (obj: TreeNode): TreeNodeData[] => {
      return Object.keys(obj)
        .map((key) => {
          const value = obj[key];
          if (value === null) {
            // 文件
            return {
              name: key,
              isDirectory: false,
              children: [],
            };
          } else {
            // 目录
            return {
              name: key,
              isDirectory: true,
              children: convertToArray(value.node),
            };
          }
        })
        .sort((a, b) => {
          // 文件夹优先
          if (a.isDirectory && !b.isDirectory) return -1;
          if (!a.isDirectory && b.isDirectory) return 1;
          // 同类型按名称排序
          return a.name.localeCompare(b.name);
        });
    };

    return convertToArray(tree);
  };

  const treeData = useMemo(
    () => (fileEntries.length > 0 ? buildTree(fileEntries) : []),
    [fileEntries]
  );

  return (
    <div className="h-full w-full flex flex-col bg-sidebar text-sidebar-foreground overflow-hidden">
      <div className="flex-1 overflow-y-auto">
        <SidebarContent>
          {uploadedFileName && (
            <SidebarGroup>
              <SidebarGroupLabel>Uploaded File</SidebarGroupLabel>
              <SidebarGroupContent>
                <SidebarMenu>
                  <SidebarMenuItem>
                    <SidebarMenuButton className="flex items-center gap-2">
                      <File className="shrink-0" />
                      <span className="truncate flex-1 min-w-0">
                        {uploadedFileName}
                      </span>
                      <div className="w-20" />
                    </SidebarMenuButton>
                    <SidebarMenuBadge className="shrink-0">
                      {fileEntries.length} files
                    </SidebarMenuBadge>
                  </SidebarMenuItem>
                </SidebarMenu>
              </SidebarGroupContent>
            </SidebarGroup>
          )}
          <SidebarGroup>
            <SidebarGroupLabel>
              {fileEntries.length > 0 ? "Archive Contents" : "Files"}
            </SidebarGroupLabel>
            <SidebarGroupContent>
              <SidebarMenu>
                {treeData.map((item, index) => (
                  <Tree
                    key={index}
                    item={item}
                    onFileSelect={onFileSelect}
                    pathPrefix=""
                  />
                ))}
              </SidebarMenu>
            </SidebarGroupContent>
          </SidebarGroup>
        </SidebarContent>
      </div>
    </div>
  );
}

function Tree({
  item,
  onFileSelect,
  pathPrefix,
}: {
  item: TreeNodeData;
  onFileSelect?: (filepath: string) => void;
  pathPrefix: string;
}) {
  const currentPath = pathPrefix ? `${pathPrefix}/${item.name}` : item.name;

  if (!item.isDirectory) {
    return (
      <SidebarMenuButton
        className="data-[active=true]:bg-transparent"
        onClick={() => onFileSelect?.(currentPath)}
      >
        <File />
        <span className="truncate">{item.name}</span>
      </SidebarMenuButton>
    );
  }

  return (
    <SidebarMenuItem>
      <Collapsible
        className="group/collapsible [&[data-state=open]>button>svg:first-child]:rotate-90"
        defaultOpen={false}
      >
        <CollapsibleTrigger asChild>
          <SidebarMenuButton>
            <ChevronRight className="transition-transform" />
            <Folder />
            <span className="truncate">{item.name}</span>
          </SidebarMenuButton>
        </CollapsibleTrigger>
        <CollapsibleContent>
          <SidebarMenuSub>
            {item.children.map((subItem, index) => (
              <Tree
                key={index}
                item={subItem}
                onFileSelect={onFileSelect}
                pathPrefix={currentPath}
              />
            ))}
          </SidebarMenuSub>
        </CollapsibleContent>
      </Collapsible>
    </SidebarMenuItem>
  );
}
