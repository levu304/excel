export type NoteText = {
  text: string;
};

export type NoteConfigs = {
  type: "note";
  note: {
    texts: NoteText[];
    margins: {
      insetmode: string;
      inset: number[];
    };
    protection: {
      locked: string;
      lockText: string;
    };
    editAs: string;
  };
};

const DEFAULT_CONFIGS: NoteConfigs = {
  type: "note",
  note: {
    texts: [],
    margins: {
      insetmode: "auto",
      inset: [0.13, 0.13, 0.25, 0.25],
    },
    protection: {
      locked: "True",
      lockText: "True",
    },
    editAs: "absolute",
  },
};

export class Note {
  private note: string | Partial<NoteConfigs["note"]>;

  constructor(note: string | Partial<NoteConfigs["note"]>) {
    this.note = note;
  }

  get model(): NoteConfigs {
    let value = null;
    switch (typeof this.note) {
      case "string":
        value = {
          texts: [
            {
              text: this.note,
            },
          ],
        };
        break;
      default:
        value = this.note;
        break;
    }
    return { type: "note", note: { ...DEFAULT_CONFIGS.note, ...value } };
  }

  set model(value: NoteConfigs) {
    const { note } = value;
    const { texts } = note;
    if (texts.length === 1 && Object.keys(texts[0]).length === 1) {
      this.note = texts[0].text;
    } else {
      this.note = note;
    }
  }
}
