<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Repository;

use App\Entity\Calculation;
use Doctrine\Bundle\DoctrineBundle\Repository\ServiceEntityRepository;
use Doctrine\Common\Persistence\ManagerRegistry;

/**
 * @method Calculation|null find($id, $lockMode = null, $lockVersion = null)
 * @method Calculation|null findOneBy(array $criteria, array $orderBy = null)
 * @method Calculation[]    findAll()
 * @method Calculation[]    findBy(array $criteria, array $orderBy = null, $limit = null, $offset = null)
 */
class CalculationRepository extends ServiceEntityRepository
{
    public function __construct(ManagerRegistry $registry)
    {
        parent::__construct($registry, Calculation::class);
    }

    // /**
    //  * @return Calculation[] Returns an array of Calculation objects
    //  */
    /*
    public function findByExampleField($value)
    {
        return $this->createQueryBuilder('c')
            ->andWhere('c.exampleField = :val')
            ->setParameter('val', $value)
            ->orderBy('c.id', 'ASC')
            ->setMaxResults(10)
            ->getQuery()
            ->getResult()
        ;
    }
    */

    /*
    public function findOneBySomeField($value): ?Calculation
    {
        return $this->createQueryBuilder('c')
            ->andWhere('c.exampleField = :val')
            ->setParameter('val', $value)
            ->getQuery()
            ->getOneOrNullResult()
        ;
    }
    */
}
