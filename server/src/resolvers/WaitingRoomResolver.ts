import { Arg, Authorized, Int, Mutation, Query, Resolver } from 'type-graphql';
import WaitingRoom, { WaitingRoomInput } from '../entity/WaitingRoom';
import dataSource from '../db';
import { ApolloError } from 'apollo-server-errors';

@Resolver(WaitingRoom)
export class WaitingRoomResolver {
  /*************************************
                    QUERY
     *************************************/

  @Query(() => [WaitingRoom])
  async getAllWaitingRooms(): Promise<WaitingRoom[]> {
    return await dataSource.getRepository(WaitingRoom).find();
  }
  /*************************************
                   MUTATION
     *************************************/

  // @Authorized<Role>(['admin'])
  @Mutation(() => WaitingRoom)
  async createWaitingRoom(
    @Arg('data') data: WaitingRoomInput
  ): Promise<WaitingRoom> {
    return await dataSource.getRepository(WaitingRoom).save(data);
  }

  @Mutation(() => WaitingRoom)
  async deleteWaitingRoom(@Arg('id', () => Int) id: number): Promise<boolean> {
    const { affected } = await dataSource.getRepository(WaitingRoom).delete(id);
    if (affected === 0)
      throw new ApolloError('Waiting room not found', 'NOT_FOUND');
    return true;
  }

  @Mutation(() => WaitingRoom)
  async updtateWilder(
    @Arg('id', () => Int) id: number,
    @Arg('data') { name }: WaitingRoomInput
  ): Promise<WaitingRoom> {
    const { affected } = await dataSource
      .getRepository(WaitingRoom)
      .update(id, { name });

    if (affected === 0) throw new ApolloError('skill not found', 'NOT_FOUND');

    return { id, name };
  }
}
