import type { LocalQueueEntry, RemoteQueueEntry } from "../../lib/helpers";

interface QueueStripProps {
  localQueue: LocalQueueEntry[];
  remoteQueue: RemoteQueueEntry[];
  selectedLocalQueueId: string | undefined;
  onQueueSelection: (entry: LocalQueueEntry) => void;
}

export function QueueStrip({
  localQueue,
  remoteQueue,
  selectedLocalQueueId,
  onQueueSelection,
}: QueueStripProps) {
  if (localQueue.length === 0 && remoteQueue.length === 0) return null;

  return (
    <section className="queue-strip">
      {localQueue.map((entry) => (
        <button
          key={entry.id}
          type="button"
          className={[
            "queue-chip",
            selectedLocalQueueId === entry.id ? "queue-chip-selected" : "",
            entry.status === "sending" ? "queue-chip-sending" : "",
          ]
            .filter(Boolean)
            .join(" ")}
          onClick={() => onQueueSelection(entry)}
        >
          <span className="queue-kind">
            {entry.status === "sending" ? "Sending next" : "Queued"}
          </span>
          <span className="queue-text">{entry.text}</span>
        </button>
      ))}

      {remoteQueue.map((entry) => (
        <div key={entry.id} className="queue-chip queue-chip-remote">
          <span className="queue-kind">
            {entry.mode === "steer" ? "Steer" : "Follow-up"}
          </span>
          <span className="queue-text">{entry.text}</span>
        </div>
      ))}
    </section>
  );
}
