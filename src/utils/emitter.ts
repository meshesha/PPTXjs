export type Unsubscribe = () => void;

/** Tiny typed event emitter (no DOM dependency). */
export class Emitter<Events extends { [K in keyof Events]: any }> {
  private handlers = new Map<keyof Events, Set<(payload: any) => void>>();

  on<K extends keyof Events>(event: K, handler: (payload: Events[K]) => void): Unsubscribe {
    let set = this.handlers.get(event);
    if (!set) this.handlers.set(event, (set = new Set()));
    set.add(handler);
    return () => this.off(event, handler);
  }

  off<K extends keyof Events>(event: K, handler: (payload: Events[K]) => void): void {
    this.handlers.get(event)?.delete(handler);
  }

  emit<K extends keyof Events>(event: K, payload: Events[K]): void {
    for (const h of [...(this.handlers.get(event) ?? [])]) {
      try {
        h(payload);
      } catch (e) {
        // a failing listener must not break rendering
        setTimeout(() => {
          throw e;
        }, 0);
      }
    }
  }

  clear(): void {
    this.handlers.clear();
  }
}
