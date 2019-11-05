import * as express from "express";
import * as debug from "debug";
import Customer from "../defs/customer";

import customers = require("../data/customers.json");


// Initialize debug logging module
const log = debug("customer");

export const customerApi = (options: any): express.Router => {
    const router = express.Router();
    log("Initializing Customer API");

    
    router.get("/customers", (req: express.Request, res: express.Response, next: express.NextFunction) => {
        let page: number;
        if (req.query.page) {
            page = parseInt(req.query.page, 10);
        } else {
            page = 0;
        }
        log(`Returning 10 customers, starting with page ${page}`);

        if (req.query.country) {
            res.json(customers
                .filter((c: Customer) => c.country === req.query.country)
                .slice(page * 10, (page + 1) * 10));
        } else {
            res.json(customers.slice(page * 10, (page + 1) * 10));
        }
    });


    router.get("/countries", (req: express.Request, res: express.Response, next: express.NextFunction) => {
        log(`Returning countries`);

        const countries = [... new Set(customers.map((c: Customer) => c.country))].sort();

        res.json(countries);
    });

    return router;
};
