import json
from engine import DocxEngine


class Repl:
    """Terminal REPL that delegates logic to DocxEngine and handles I/O."""

    def __init__(self):
        self.engine = None

    def print_help(self):
        print("\nCommands:")
        print("  load [filename]              - Load a .docx file")
        print("  map                          - Show document structure JSON")
        print("  replace [id] [text...]       - Replace text in ID")
        print("  insert_after [id] [text...]  - Insert paragraph after ID")
        print("  delete [id]                  - Delete element")
        print("  format [id] [prop] [val]     - Set prop (bold/italic/size) to value")
        print("  save [filename]              - Save output")
        print("  validate [filename]          - Check integrity\n")

    def run(self):
        print("--- Precision Document Editor Prototype ---")
        print("Type 'help' for commands or 'exit' to quit.")

        while True:
            try:
                parts = input("> ").strip().split(" ")
                cmd = parts[0].lower() if parts else ""
                args = parts[1:]

                if cmd == "exit":
                    break

                elif cmd == "help":
                    self.print_help()

                elif cmd == "load":
                    if not args:
                        print("Usage: load [filename]")
                        continue
                    try:
                        self.engine = DocxEngine(args[0])
                        print(f"Loaded {args[0]}")
                        print(f"Stats: {self.engine.structure_cache['metadata']}")
                    except Exception as e:
                        print(f"Error loading: {e}")

                elif cmd == "map":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    # Return dict from engine, pretty-print JSON here
                    data = self.engine.map_structure()
                    print(json.dumps(data, indent=2))

                elif cmd == "replace":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    if len(args) < 2:
                        print("Usage: replace [id] [new text]")
                        continue
                    tgt_id = args[0]
                    new_text = " ".join(args[1:]).strip('"').strip("'")
                    print(self.engine.replace_text(tgt_id, new_text))

                elif cmd == "insert_after":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    if len(args) < 2:
                        print("Usage: insert_after [id] [new text]")
                        continue
                    tgt_id = args[0]
                    new_text = " ".join(args[1:]).strip('"').strip("'")
                    print(self.engine.insert_after(tgt_id, new_text))

                elif cmd == "delete":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    if len(args) < 1:
                        print("Usage: delete [id]")
                        continue
                    print(self.engine.delete_element(args[0]))

                elif cmd == "format":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    if len(args) < 3:
                        print("Usage: format [id] [prop] [value]")
                        continue
                    print(self.engine.format_element(args[0], args[1], args[2]))

                elif cmd == "save":
                    if not self.engine:
                        print("No document loaded.")
                        continue
                    if len(args) < 1:
                        print("Usage: save [filename]")
                        continue
                    print(self.engine.save(args[0]))

                elif cmd == "validate":
                    if len(args) < 1:
                        print("Usage: validate [filename]")
                        continue
                    # Static validation, independent of loaded engine
                    print(DocxEngine.validate(args[0]))

                else:
                    print("Unknown command.")

            except Exception as e:
                print(f"An error occurred: {e}")


def main():
    Repl().run()


if __name__ == "__main__":
    main()