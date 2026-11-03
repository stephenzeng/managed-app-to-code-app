import { useMemo, useState, type FormEvent } from "react";
import { useMutation, useQuery, useQueryClient } from "@tanstack/react-query";
import {
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  MoreHorizontal,
  Plus,
  Search,
  Trash2,
  Users,
} from "lucide-react";
import { toast } from "sonner";
import { Issue_tracker_listService } from "@/generated";
import type {
  Issue_tracker_listRead,
  Issue_tracker_listWrite,
} from "@/generated/models/Issue_tracker_listModel";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Calendar } from "@/components/ui/calendar";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import {
  Sheet,
  SheetContent,
  SheetDescription,
  SheetFooter,
  SheetHeader,
  SheetTitle,
} from "@/components/ui/sheet";
import { Skeleton } from "@/components/ui/skeleton";
import { ConnectorErrorState, EmptyState } from "@/components/states";
import { useConfirm } from "@/hooks/use-confirm";
import { cn } from "@/lib/utils";

type Issue = Issue_tracker_listRead;
type IssueWrite = Partial<Omit<Issue_tracker_listWrite, "ID">>;
type ChoiceField = "Status" | "Priority" | "Category";

const PAGE_SIZE = 100;
const NONE = "__none__";
const ALL = "__all__";
const CONNECTOR = "SharePoint";

/* ---------------- tone map (status / priority) ---------------- */

type Tone = "green" | "amber" | "red" | "blue" | "gray";

const TONE_BADGE: Record<Tone, string> = {
  green: "bg-green-100 text-green-800 dark:bg-green-950 dark:text-green-200",
  amber: "bg-amber-100 text-amber-900 dark:bg-amber-950 dark:text-amber-200",
  red: "bg-red-100 text-red-800 dark:bg-red-950 dark:text-red-200",
  blue: "bg-blue-100 text-blue-800 dark:bg-blue-950 dark:text-blue-200",
  gray: "bg-muted text-foreground",
};

const TONE_DOT: Record<Tone, string> = {
  green: "bg-green-600",
  amber: "bg-amber-500",
  red: "bg-red-600",
  blue: "bg-blue-600",
  gray: "bg-zinc-400",
};

function statusTone(value?: string): Tone {
  const v = (value ?? "").toLowerCase();
  if (!v) return "gray";
  if (/(resolv|clos|done|complet|fixed)/.test(v)) return "green";
  if (/(block|reject|cancel|fail)/.test(v)) return "red";
  if (/(progress|active|working|review|testing)/.test(v)) return "blue";
  if (/(new|open|pending|wait|hold|triage|backlog|not started)/.test(v)) return "amber";
  return "gray";
}

function priorityTone(value?: string): Tone {
  const v = (value ?? "").toLowerCase();
  if (/(critical|urgent|highest|high|p0|p1)/.test(v)) return "red";
  if (/(medium|normal|moderate|p2)/.test(v)) return "amber";
  if (/(low|minor|p3|p4)/.test(v)) return "blue";
  return "gray";
}

function ToneBadge({ tone, label, prefix }: { tone: Tone; label: string; prefix?: string }) {
  return (
    <span
      className={cn(
        "inline-flex items-center gap-1.5 rounded-full px-2 py-0.5 text-xs font-medium forced-colors:border",
        TONE_BADGE[tone],
      )}
    >
      <span className={cn("size-1.5 rounded-full forced-colors:bg-[CanvasText]", TONE_DOT[tone])} aria-hidden="true" />
      {prefix && <span className="sr-only">{prefix}: </span>}
      {label}
    </span>
  );
}

/* ---------------- helpers ---------------- */

function stripHtml(value?: string): string {
  if (!value) return "";
  if (!/[<&]/.test(value)) return value;
  const doc = new DOMParser().parseFromString(value, "text/html");
  return (doc.body.textContent ?? "").trim();
}

function toLocalDate(value?: string): Date | undefined {
  if (!value) return undefined;
  if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [y, m, d] = value.split("-").map(Number);
    return new Date(y, m - 1, d);
  }
  const d = new Date(value);
  return Number.isNaN(d.getTime()) ? undefined : d;
}

function toYmd(date?: Date): string {
  if (!date) return "";
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${date.getFullYear()}-${m}-${d}`;
}

const dateFmt = new Intl.DateTimeFormat(undefined, { day: "numeric", month: "short", year: "numeric" });

function formatDate(value?: string): string {
  const d = toLocalDate(value);
  return d ? dateFmt.format(d) : "";
}

function relativeDue(value?: string): string {
  const d = toLocalDate(value);
  if (!d) return "";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const due = new Date(d);
  due.setHours(0, 0, 0, 0);
  const days = Math.round((due.getTime() - today.getTime()) / 86_400_000);
  if (days === 0) return "today";
  if (days === 1) return "tomorrow";
  if (days === -1) return "yesterday";
  return days > 0 ? `in ${days} days` : `${-days} days ago`;
}

/* ---------------- data hooks ---------------- */

function parseChoices(payload: unknown): string[] {
  const raw = Array.isArray(payload)
    ? payload
    : typeof payload === "object" && payload !== null && "value" in payload && Array.isArray((payload as { value: unknown }).value)
      ? (payload as { value: unknown[] }).value
      : null;
  if (raw === null) throw new Error("Unsupported SharePoint Choice response shape");
  return raw.map((item) => {
    if (typeof item === "string") return item;
    if (typeof item === "object" && item !== null && "Value" in item && typeof (item as { Value: unknown }).Value === "string") {
      return (item as { Value: string }).Value;
    }
    throw new Error("Unsupported SharePoint Choice item shape");
  });
}

function useChoices(field: ChoiceField) {
  return useQuery({
    queryKey: ["choices", field],
    staleTime: 5 * 60_000,
    queryFn: async () => {
      const result = await Issue_tracker_listService.getReferencedEntity("", field);
      if (!result.success) {
        console.error("connector error", result.error);
        throw result.error;
      }
      return parseChoices(result.data);
    },
  });
}

function useIssueMutations() {
  const qc = useQueryClient();
  const refresh = () => qc.invalidateQueries({ queryKey: ["issues"] });

  const save = useMutation({
    mutationFn: async ({ id, fields }: { id?: number; fields: IssueWrite }) => {
      const result =
        id === undefined
          ? await Issue_tracker_listService.create(fields)
          : await Issue_tracker_listService.update(String(id), fields);
      if (!result.success) {
        console.error("connector error", result.error);
        throw result.error;
      }
      return result.data;
    },
    onSuccess: (_d, vars) => {
      toast.success(vars.id === undefined ? "Issue created" : "Issue updated");
      void refresh();
    },
    onError: () => toast.error("Couldn't save the issue. Try again."),
  });

  const remove = useMutation({
    mutationFn: async (id: number) => {
      await Issue_tracker_listService.delete(String(id));
    },
    onSuccess: () => {
      toast.success("Issue deleted");
      void refresh();
    },
    onError: (error) => {
      console.error("connector error", error);
      toast.error("Couldn't delete the issue. Try again.");
    },
  });

  return { save, remove };
}

/* ---------------- page ---------------- */

export function HomePage() {
  const statuses = useChoices("Status");
  const priorities = useChoices("Priority");
  const categories = useChoices("Category");

  const [skipToken, setSkipToken] = useState<string | undefined>();
  const [previousTokens, setPreviousTokens] = useState<(string | undefined)[]>([]);
  const [search, setSearch] = useState("");
  const [priorityFilter, setPriorityFilter] = useState(ALL);
  const [categoryFilter, setCategoryFilter] = useState(ALL);
  const [editing, setEditing] = useState<Issue | "new" | null>(null);

  const issues = useQuery({
    queryKey: ["issues", { skipToken }],
    queryFn: async () => {
      const result = await Issue_tracker_listService.getAll({
        orderBy: ["Modified desc"],
        top: PAGE_SIZE,
        skipToken,
      });
      if (!result.success) {
        console.error("connector error", result.error);
        throw result.error;
      }
      return result;
    },
  });

  const { save, remove } = useIssueMutations();
  const confirm = useConfirm();

  const rows = useMemo(() => issues.data?.data ?? [], [issues.data]);
  const nextToken = issues.data?.skipToken?.trim() || undefined;
  const pageNumber = previousTokens.length + 1;

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return rows.filter((r) => {
      if (priorityFilter !== ALL && (r.Priority?.Value ?? NONE) !== priorityFilter) return false;
      if (categoryFilter !== ALL && (r.Category?.Value ?? NONE) !== categoryFilter) return false;
      if (!q) return true;
      return [r.Title, stripHtml(r.IssueDescription), r.AssignedTeam, r.Category?.Value]
        .filter(Boolean)
        .some((t) => String(t).toLowerCase().includes(q));
    });
  }, [rows, search, priorityFilter, categoryFilter]);

  const lanes = useMemo(() => {
    const configured = statuses.data ?? [];
    const seen = new Set(configured);
    const extra: string[] = [];
    for (const r of rows) {
      const v = r.Status?.Value;
      if (v && !seen.has(v)) {
        seen.add(v);
        extra.push(v);
      }
    }
    const list = [...configured, ...extra].map((s) => ({ key: s, label: s }));
    if (rows.some((r) => !r.Status?.Value)) list.push({ key: NONE, label: "No status" });
    return list.map((lane) => ({
      ...lane,
      items: filtered.filter((r) => (r.Status?.Value ?? NONE) === lane.key),
    }));
  }, [statuses.data, rows, filtered]);

  function moveTo(issue: Issue, status: string) {
    if (issue.ID === undefined || issue.Status?.Value === status) return;
    save.mutate({ id: issue.ID, fields: { Status: status } });
  }

  async function deleteIssue(issue: Issue) {
    if (issue.ID === undefined) return;
    const ok = await confirm({
      title: `Delete "${issue.Title ?? "Untitled issue"}"?`,
      description: "This removes the item from the SharePoint list.",
      confirmLabel: "Delete",
      destructive: true,
    });
    if (!ok) return;
    remove.mutate(issue.ID, { onSuccess: () => setEditing(null) });
  }

  const filtersActive = search.trim() !== "" || priorityFilter !== ALL || categoryFilter !== ALL;

  return (
    <div className="mx-auto flex h-full w-full max-w-[1600px] flex-col gap-4 px-4 py-5 sm:px-6">
      {/* toolbar */}
      <div className="flex flex-wrap items-end gap-3">
        <div className="mr-auto">
          <h1 className="text-xl font-semibold tracking-tight">Issues</h1>
          {issues.isSuccess && (
            <p className="text-sm text-muted-foreground" aria-live="polite">
              {filtersActive ? `${filtered.length} of ${rows.length}` : rows.length} shown
              {(pageNumber > 1 || nextToken) && ` · page ${pageNumber}`}
            </p>
          )}
        </div>
        <div className="relative">
          <Search className="pointer-events-none absolute left-2.5 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" aria-hidden="true" />
          <Input
            type="search"
            aria-label="Search issues"
            placeholder="Search issues"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="w-56 bg-background pl-8"
          />
        </div>
        <FilterSelect
          label="Priority"
          allLabel="All priorities"
          value={priorityFilter}
          onChange={setPriorityFilter}
          options={priorities.data ?? []}
        />
        <FilterSelect
          label="Category"
          allLabel="All categories"
          value={categoryFilter}
          onChange={setCategoryFilter}
          options={categories.data ?? []}
        />
        <Button onClick={() => setEditing("new")}>
          <Plus aria-hidden="true" />
          New issue
        </Button>
      </div>

      {/* board */}
      <section id="board" tabIndex={-1} aria-label="Issue board" className="flex-1 outline-none">
        {issues.isPending || statuses.isPending ? (
          <BoardSkeleton />
        ) : issues.isError ? (
          <ConnectorErrorState
            error={issues.error}
            connector={CONNECTOR}
            action={<Button variant="outline" onClick={() => void issues.refetch()}>Retry</Button>}
          />
        ) : statuses.isError ? (
          <ConnectorErrorState
            error={statuses.error}
            connector={CONNECTOR}
            action={<Button variant="outline" onClick={() => void statuses.refetch()}>Retry</Button>}
          />
        ) : previousTokens.length === 0 && rows.length === 0 && !nextToken ? (
          <EmptyState
            title="No issues yet"
            action={
              <Button onClick={() => setEditing("new")}>
                <Plus aria-hidden="true" />
                New issue
              </Button>
            }
          />
        ) : (
          <div className="flex gap-4 overflow-x-auto pb-3">
            {lanes.map((lane) => (
              <section
                key={lane.key}
                aria-labelledby={`lane-${lane.key}`}
                className="flex w-72 shrink-0 flex-col rounded-lg border border-border bg-background"
              >
                <header className="flex items-center gap-2 border-b border-border px-3 py-2.5">
                  <span className={cn("size-2.5 rounded-full forced-colors:bg-[CanvasText]", TONE_DOT[statusTone(lane.key === NONE ? "" : lane.label)])} aria-hidden="true" />
                  <h2 id={`lane-${lane.key}`} className="text-sm font-semibold">
                    {lane.label}
                  </h2>
                  <span className="ml-auto rounded-full bg-muted px-2 text-xs font-medium tabular-nums">
                    <span className="sr-only">Issues: </span>
                    {lane.items.length}
                  </span>
                </header>
                <ul className="flex min-h-24 flex-col gap-2 p-2">
                  {lane.items.length === 0 && (
                    <li className="px-2 py-6 text-center text-sm text-muted-foreground">None</li>
                  )}
                  {lane.items.map((issue) => (
                    <li key={issue.ID}>
                      <IssueCard
                        issue={issue}
                        statuses={statuses.data ?? []}
                        onOpen={() => setEditing(issue)}
                        onMove={(s) => moveTo(issue, s)}
                        onDelete={() => void deleteIssue(issue)}
                        busy={save.isPending && save.variables?.id === issue.ID}
                      />
                    </li>
                  ))}
                </ul>
              </section>
            ))}
          </div>
        )}

        {issues.isSuccess && (previousTokens.length > 0 || nextToken) && (
          <nav aria-label="Pages" className="mt-2 flex items-center justify-end gap-2">
            <Button
              variant="outline"
              size="sm"
              disabled={previousTokens.length === 0 || issues.isFetching}
              onClick={() => {
                setSkipToken(previousTokens[previousTokens.length - 1]);
                setPreviousTokens((t) => t.slice(0, -1));
              }}
            >
              <ChevronLeft aria-hidden="true" />
              Previous page
            </Button>
            <Button
              variant="outline"
              size="sm"
              disabled={!nextToken || issues.isFetching}
              onClick={() => {
                if (!nextToken) return;
                setPreviousTokens((t) => [...t, skipToken]);
                setSkipToken(nextToken);
              }}
            >
              {issues.isFetching ? "Loading…" : "Next page"}
              <ChevronRight aria-hidden="true" />
            </Button>
          </nav>
        )}
      </section>

      <IssueSheet
        key={editing === "new" ? "new" : editing?.ID ?? "closed"}
        issue={editing}
        statuses={statuses.data ?? []}
        priorities={priorities.data ?? []}
        categories={categories.data ?? []}
        saving={save.isPending}
        deleting={remove.isPending}
        onClose={() => setEditing(null)}
        onSave={(fields) =>
          save.mutate(
            { id: editing === "new" || editing === null ? undefined : editing.ID, fields },
            { onSuccess: () => setEditing(null) },
          )
        }
        onDelete={() => editing && editing !== "new" && void deleteIssue(editing)}
      />
    </div>
  );
}

/* ---------------- pieces ---------------- */

function FilterSelect({
  label,
  allLabel,
  value,
  onChange,
  options,
}: {
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: string[];
  allLabel: string;
}) {
  return (
    <Select value={value} onValueChange={onChange}>
      <SelectTrigger aria-label={`Filter by ${label.toLowerCase()}`} className="w-40 bg-background">
        <SelectValue />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value={ALL}>{allLabel}</SelectItem>
        {options.map((o) => (
          <SelectItem key={o} value={o}>
            {o}
          </SelectItem>
        ))}
        <SelectItem value={NONE}>No {label.toLowerCase()}</SelectItem>
      </SelectContent>
    </Select>
  );
}

function IssueCard({
  issue,
  statuses,
  onOpen,
  onMove,
  onDelete,
  busy,
}: {
  issue: Issue;
  statuses: string[];
  onOpen: () => void;
  onMove: (status: string) => void;
  onDelete: () => void;
  busy: boolean;
}) {
  const description = stripHtml(issue.IssueDescription);
  const due = issue.DueDate;
  return (
    <article
      className={cn(
        "group relative rounded-md border border-border bg-card p-3 shadow-xs transition-shadow hover:shadow-sm focus-within:ring-2 focus-within:ring-ring",
        busy && "opacity-60",
      )}
      aria-busy={busy}
    >
      <div className="flex items-start gap-2">
        <h3 className="flex-1 text-sm font-medium leading-snug">
          <button
            type="button"
            onClick={onOpen}
            aria-haspopup="dialog"
            className="text-left outline-none after:absolute after:inset-0 after:content-['']"
          >
            {issue.Title || "Untitled issue"}
          </button>
        </h3>
        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              variant="ghost"
              size="icon"
              className="relative z-10 -mr-1 -mt-1 size-7"
              aria-label={`Actions for ${issue.Title || "untitled issue"}`}
            >
              <MoreHorizontal aria-hidden="true" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="end">
            <DropdownMenuLabel>Move to</DropdownMenuLabel>
            {statuses.map((s) => (
              <DropdownMenuItem
                key={s}
                disabled={issue.Status?.Value === s}
                onSelect={() => onMove(s)}
              >
                <span className={cn("size-2 rounded-full", TONE_DOT[statusTone(s)])} aria-hidden="true" />
                {s}
              </DropdownMenuItem>
            ))}
            <DropdownMenuSeparator />
            <DropdownMenuItem variant="destructive" onSelect={onDelete}>
              <Trash2 aria-hidden="true" />
              Delete
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>
      </div>
      {description && (
        <p className="mt-1 line-clamp-2 text-xs text-muted-foreground">{description}</p>
      )}
      <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
        {issue.Priority?.Value && (
          <ToneBadge tone={priorityTone(issue.Priority.Value)} label={issue.Priority.Value} prefix="Priority" />
        )}
        {issue.Category?.Value && (
          <span className="rounded-full border border-border px-2 py-0.5 text-xs">
            <span className="sr-only">Category: </span>
            {issue.Category.Value}
          </span>
        )}
      </div>
      {(issue.AssignedTeam || due) && (
        <div className="mt-2 flex flex-wrap items-center gap-x-3 gap-y-1 text-xs text-muted-foreground">
          {issue.AssignedTeam && (
            <span className="inline-flex items-center gap-1">
              <Users className="size-3.5" aria-hidden="true" />
              <span className="sr-only">Team: </span>
              {issue.AssignedTeam}
            </span>
          )}
          {due && (
            <span className="inline-flex items-center gap-1" title={formatDate(due)}>
              <CalendarDays className="size-3.5" aria-hidden="true" />
              Due {formatDate(due)} ({relativeDue(due)})
            </span>
          )}
        </div>
      )}
    </article>
  );
}

function DateField({
  id,
  label,
  value,
  onChange,
}: {
  id: string;
  label: string;
  value: string;
  onChange: (v: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const selected = toLocalDate(value);
  return (
    <div className="grid gap-1.5">
      <Label htmlFor={id}>{label}</Label>
      <div className="flex gap-1">
        <Popover open={open} onOpenChange={setOpen}>
          <PopoverTrigger asChild>
            <Button id={id} variant="outline" className="flex-1 justify-start font-normal">
              <CalendarDays aria-hidden="true" />
              {selected ? dateFmt.format(selected) : <span className="text-muted-foreground">Pick a date</span>}
            </Button>
          </PopoverTrigger>
          <PopoverContent className="w-auto p-0" align="start">
            <Calendar
              mode="single"
              selected={selected}
              defaultMonth={selected}
              onSelect={(d) => {
                onChange(toYmd(d));
                setOpen(false);
              }}
            />
          </PopoverContent>
        </Popover>
        {value && (
          <Button variant="ghost" size="sm" onClick={() => onChange("")} aria-label={`Clear ${label.toLowerCase()}`}>
            Clear
          </Button>
        )}
      </div>
    </div>
  );
}

function ChoiceField({
  id,
  label,
  value,
  onChange,
  options,
}: {
  id: string;
  label: string;
  value: string;
  onChange: (v: string) => void;
  options: string[];
}) {
  const all = value && !options.includes(value) ? [value, ...options] : options;
  return (
    <div className="grid gap-1.5">
      <Label htmlFor={id}>{label}</Label>
      <Select value={value || NONE} onValueChange={(v) => onChange(v === NONE ? "" : v)}>
        <SelectTrigger id={id} className="w-full">
          <SelectValue />
        </SelectTrigger>
        <SelectContent>
          <SelectItem value={NONE}>None</SelectItem>
          {all.map((o) => (
            <SelectItem key={o} value={o}>
              {o}
            </SelectItem>
          ))}
        </SelectContent>
      </Select>
    </div>
  );
}

function IssueSheet({
  issue,
  statuses,
  priorities,
  categories,
  saving,
  deleting,
  onClose,
  onSave,
  onDelete,
}: {
  issue: Issue | "new" | null;
  statuses: string[];
  priorities: string[];
  categories: string[];
  saving: boolean;
  deleting: boolean;
  onClose: () => void;
  onSave: (fields: IssueWrite) => void;
  onDelete: () => void;
}) {
  const existing = issue && issue !== "new" ? issue : null;
  const [title, setTitle] = useState(existing?.Title ?? "");
  const [description, setDescription] = useState(stripHtml(existing?.IssueDescription));
  const [status, setStatus] = useState(existing?.Status?.Value ?? (issue === "new" ? statuses[0] ?? "" : ""));
  const [priority, setPriority] = useState(existing?.Priority?.Value ?? "");
  const [category, setCategory] = useState(existing?.Category?.Value ?? "");
  const [team, setTeam] = useState(existing?.AssignedTeam ?? "");
  const [reported, setReported] = useState(
    existing ? toYmd(toLocalDate(existing.ReportedDate)) : issue === "new" ? toYmd(new Date()) : "",
  );
  const [due, setDue] = useState(toYmd(toLocalDate(existing?.DueDate)));
  const [touched, setTouched] = useState(false);

  const titleError = touched && !title.trim() ? "Enter a title." : "";

  function submit(e: FormEvent) {
    e.preventDefault();
    setTouched(true);
    if (!title.trim()) return;
    const fields: IssueWrite = {
      Title: title.trim(),
      IssueDescription: description,
      AssignedTeam: team,
    };
    if (status) fields.Status = status;
    if (priority) fields.Priority = priority;
    if (category) fields.Category = category;
    if (reported) fields.ReportedDate = reported;
    if (due) fields.DueDate = due;
    onSave(fields);
  }

  return (
    <Sheet open={issue !== null} onOpenChange={(o) => !o && onClose()}>
      <SheetContent className="flex w-full flex-col gap-0 sm:max-w-md">
        <form onSubmit={submit} className="flex h-full flex-col" noValidate>
          <SheetHeader className="border-b border-border">
            <SheetTitle>{existing ? existing.Title || "Untitled issue" : "New issue"}</SheetTitle>
            {existing ? (
              <SheetDescription>
                #{existing.ID}
                {existing.Author?.DisplayName && ` · reported by ${existing.Author.DisplayName}`}
                {existing.Modified && ` · updated ${formatDate(existing.Modified)}`}
              </SheetDescription>
            ) : (
              <SheetDescription className="sr-only">Create an issue</SheetDescription>
            )}
          </SheetHeader>
          <div className="flex-1 space-y-4 overflow-y-auto p-4">
            <div className="grid gap-1.5">
              <Label htmlFor="f-title">Title</Label>
              <Input
                id="f-title"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                onBlur={() => setTouched(true)}
                aria-invalid={!!titleError}
                aria-describedby={titleError ? "f-title-err" : undefined}
                required
              />
              {titleError && (
                <p id="f-title-err" className="text-sm text-destructive">
                  {titleError}
                </p>
              )}
            </div>
            <div className="grid gap-1.5">
              <Label htmlFor="f-desc">Description</Label>
              <Textarea id="f-desc" rows={5} value={description} onChange={(e) => setDescription(e.target.value)} />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <ChoiceField id="f-status" label="Status" value={status} onChange={setStatus} options={statuses} />
              <ChoiceField id="f-priority" label="Priority" value={priority} onChange={setPriority} options={priorities} />
              <ChoiceField id="f-category" label="Category" value={category} onChange={setCategory} options={categories} />
              <div className="grid gap-1.5">
                <Label htmlFor="f-team">Assigned team</Label>
                <Input id="f-team" value={team} onChange={(e) => setTeam(e.target.value)} />
              </div>
            </div>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <DateField id="f-reported" label="Reported" value={reported} onChange={setReported} />
              <DateField id="f-due" label="Due" value={due} onChange={setDue} />
            </div>
          </div>
          <SheetFooter className="flex-row border-t border-border">
            {existing && (
              <Button type="button" variant="ghost" className="mr-auto text-destructive" onClick={onDelete} disabled={deleting}>
                <Trash2 aria-hidden="true" />
                Delete
              </Button>
            )}
            <Button type="button" variant="outline" onClick={onClose} className={cn(!existing && "ml-auto")}>
              Cancel
            </Button>
            <Button type="submit" disabled={saving}>
              {saving ? "Saving…" : existing ? "Save changes" : "Create issue"}
            </Button>
          </SheetFooter>
        </form>
      </SheetContent>
    </Sheet>
  );
}

function BoardSkeleton() {
  return (
    <div role="status" aria-label="Loading issues" className="flex gap-4 overflow-hidden">
      {Array.from({ length: 4 }).map((_, i) => (
        <div key={i} className="w-72 shrink-0 space-y-2 rounded-lg border border-border bg-background p-2">
          <Skeleton className="h-6 w-28" />
          <Skeleton className="h-24 w-full" />
          <Skeleton className="h-24 w-full" />
        </div>
      ))}
    </div>
  );
}
