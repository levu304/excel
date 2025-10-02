export class DataValidations {
  private model: Record<string, unknown>;

  constructor(model?: Record<string, unknown>) {
    this.model = model || {};
  }

  add(address: string, validation: unknown) {
    return (this.model[address] = validation);
  }

  find(address: string) {
    return this.model[address];
  }

  remove(address: string) {
    this.model[address] = undefined;
  }
}
