import { initializeBlock, useBase } from "@airtable/blocks/interface/ui";
import "./style.css";

function HelloWorldApp() {
    const base = useBase();

    return (
        <div className="flex items-center justify-center min-h-screen w-full bg-gray-50 dark:bg-gray-950 p-6">
            <div className="bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-800 rounded-2xl p-10 max-w-md w-full shadow-sm">
                <span className="inline-block text-xs font-semibold uppercase tracking-widest text-gray-400 bg-gray-100 dark:bg-gray-800 rounded-md px-2.5 py-1 mb-5">
                    Executive Dashboard
                </span>
                <h1 className="text-2xl font-bold text-gray-900 dark:text-white tracking-tight mb-2">
                    Hello, {base.name}
                </h1>
                <p className="text-sm text-gray-400 mb-6">
                    Block ID:{" "}
                    <code className="font-mono text-xs bg-gray-100 dark:bg-gray-800 text-purple-600 dark:text-purple-400 rounded px-1.5 py-0.5">
                        blkMvLVaFtkU0Jcyk
                    </code>
                </p>
                <div className="border-t border-gray-100 dark:border-gray-800 mb-6" />
                <p className="text-sm text-gray-500 dark:text-gray-400 leading-relaxed">
                    Your custom extension is live and connected to this base.
                    Start building your executive dashboard here.
                </p>
            </div>
        </div>
    );
}

initializeBlock({interface: () => <HelloWorldApp />});
