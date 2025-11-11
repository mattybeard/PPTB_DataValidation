import React from "react";

export class ServiceProvider {
  private services = new Map();

  get<T>(serviceName: string): T {
    if (this.services.has(serviceName)) {
      return this.services.get(serviceName) as T;
    } else throw new Error(`Service ${serviceName} not registered in ServiceProvider`);
  }
  register<T>(serviceName: string, service: T): void {
    this.services.set(serviceName, service);
  }
}

export const ServiceProviderContext = React.createContext<ServiceProvider>(new ServiceProvider());
