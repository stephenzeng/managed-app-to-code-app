import { Outlet, useLocation } from "react-router-dom";
import { motion } from "motion/react";
import { Bug } from "lucide-react";
import { Toaster } from "@/components/ui/sonner";

export function AppShell() {
  const { pathname } = useLocation();
  return (
    <div className="flex min-h-screen flex-col bg-muted/40">
      <a
        href="#board"
        className="sr-only focus:not-sr-only focus:fixed focus:left-3 focus:top-3 focus:z-50 focus:rounded-md focus:bg-background focus:px-3 focus:py-2 focus:text-sm focus:shadow focus:ring-2 focus:ring-ring"
      >
        Skip to issue board
      </a>
      <header className="sticky top-0 z-30 border-b border-border bg-background/95 backdrop-blur">
        <div className="mx-auto flex h-14 w-full max-w-[1600px] items-center gap-3 px-4 sm:px-6">
          <span
            className="flex size-8 items-center justify-center rounded-md bg-primary text-primary-foreground"
            aria-hidden="true"
          >
            <Bug className="size-4" />
          </span>
          <span className="text-base font-semibold tracking-tight">Kooka Issue Tracker</span>
        </div>
      </header>
      <main className="flex flex-1 flex-col">
        <motion.div
          key={pathname}
          initial={{ opacity: 0, y: 6 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.15, ease: "easeOut" }}
          className="flex-1"
        >
          <Outlet />
        </motion.div>
      </main>
      <Toaster />
    </div>
  );
}

export default AppShell;
