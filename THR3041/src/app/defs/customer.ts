export default class Customer {

    public id: number;
    public firstName: string;
    public lastName: string;
    public email: string;
    public avatar: string;
    public country: string;

    constructor() { }

    public name(): string {
        return this.firstName + " " + this.lastName;
    }
}
