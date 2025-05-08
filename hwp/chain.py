from .app import HService
from typing import List


def chain(*services: HService) -> HService:
    """
    Chains multiple HServices together.

    Args:
        *services: A variable number of HService instances.

    Returns:
        ChainedHService: A new HService that executes all the provided services in order.
    """
    return ChainedHService(services)


class ChainedHService(HService):
    def __init__(self, services: List[HService]):
        self.services = services

    def execute(self, com) -> tuple:
        """
        Returns:
            - results: A tuple of results from each service in the chain.
                For example, if there are two services in the chain,
                the result will be ((..result1), (..result2)).
        """
        results = []

        for service in self.services:
            results.append(service.execute(com))

        return tuple(results)
