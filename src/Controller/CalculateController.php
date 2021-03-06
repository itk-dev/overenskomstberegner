<?php

/*
 * This file is part of itk-dev/overenskomstberegner.
 *
 * (c) 2019 ITK Development
 *
 * This source file is subject to the MIT license.
 */

namespace App\Controller;

use App\Calculator\Manager;
use App\Entity\Calculation;
use Exception;
use PhpOffice\PhpSpreadsheet\Writer\Html;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\Form\Extension\Core\Type\CheckboxType;
use Symfony\Component\Form\Extension\Core\Type\DateType;
use Symfony\Component\Form\Extension\Core\Type\FileType;
use Symfony\Component\Form\Extension\Core\Type\IntegerType;
use Symfony\Component\Form\Extension\Core\Type\SubmitType;
use Symfony\Component\Form\Extension\Core\Type\TextType;
use Symfony\Component\HttpFoundation\File\UploadedFile;
use Symfony\Component\HttpFoundation\Request;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\StreamedResponse;
use Symfony\Component\Routing\Annotation\Route;
use Symfony\Component\Validator\Constraints\File;

/**
 * Class CalculateController.
 *
 * @Route("/calculate", name="calculate_")
 */
class CalculateController extends AbstractController
{
    /** @var Manager */
    private $manager;

    public function __construct(Manager $manager)
    {
        $this->manager = $manager;
    }

    /**
     * @Route("/show/{id}", name="show")
     *
     * @param Request     $request
     * @param Calculation $calculation
     *
     * @return Response
     */
    public function show(Request $request, Calculation $calculation)
    {
        $calculator = $this->manager->createCalculator($calculation->getCalculator(), $calculation->getCalculatorSettings());
        $formulas = $this->manager->getFormulas($calculator);

        return $this->render('calculation/show.html.twig', [
            'calculation' => $calculation,
            'formulas' => $formulas,
        ]);
    }

    /**
     * @Route("/run/{id}", name="run")
     *
     * @param Request     $request
     * @param Calculation $calculation
     *
     * @return Response
     */
    public function run(Request $request, Calculation $calculation)
    {
        $form = $this->buildForm($calculation);

        $form->handleRequest($request);
        $result = null;

        if ($form->isSubmitted()) {
            if ($form->isValid()) {
                $arguments = $form->getData();
                $preview = $form->get('preview')->isClicked();
                /** @var UploadedFile $file */
                $file = $form->get('input')->getData();
                // this is needed to safely include the file name as part of the URL
                $safeFilename = transliterator_transliterate('Any-Latin; Latin-ASCII; [^A-Za-z0-9_] remove; Lower()', $file->getClientOriginalName());
                $newFilename = $safeFilename.'-'.uniqid('', false).'.'.$file->getClientOriginalExtension();
                $targetPathname = sys_get_temp_dir().'/'.$newFilename;
                $file->move(
                    \dirname($targetPathname),
                    basename($targetPathname)
                );
                $input = $targetPathname;

                try {
                    $result = $this->manager->calculate(
                        $calculation->getCalculator(),
                        $calculation->getCalculatorSettings(),
                        $arguments,
                        $input
                    );

                    $filename = transliterator_transliterate('Any-Latin; Latin-ASCII; [^A-Za-z0-9_] remove; Lower()', $calculation->getName())
                        .'-'.(new \DateTime())->format('Y-m-d\TH.i.s')
                        .'.xlsx';

                    if ($preview) {
                        $writer = new Html($result);

                        return $this->render('calculate/preview.html.twig', [
                            'content' => $writer->generateSheetData(),
                            'title' => $filename,
                        ]);
                    }

                    $contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
                    $writer = new Xlsx($result);
                    $response = new StreamedResponse(
                        static function () use ($writer) {
                            $writer->save('php://output');
                        }
                    );

                    $response->headers->set('content-type', $contentType);
                    $response->headers->set('content-disposition', 'attachment; filename="'.$filename.'"');
                    $response->headers->set('cache-control', 'max-age=0');

                    return $response;
                } catch (Exception $exception) {
                    return $this->render('calculation/exception.html.twig', [
                        'exception' => $exception,
                    ]);
                }
            }
        }

        return $this->render('calculation/run.html.twig', [
            'calculation' => $calculation,
            'form' => $form->createView(),
            'result' => $result,
        ]);
    }

    private function buildForm(Calculation $calculation)
    {
        $builder = $this->createFormBuilder()
            ->add('input', FileType::class, [
                'constraints' => [
                    new File([
                        'maxSize' => '1024k',
                        //                        'mimeTypes' => [
                        //                            'text/csv',
                        //                            'application/vnd.ms-excel',
                        //                        ],
                        'mimeTypesMessage' => 'Please upload a valid CSV or Excel document',
                    ]),
                ],
            ]);

        $calculator = $this->manager->createCalculator($calculation->getCalculator(), $calculation->getCalculatorSettings());
        foreach ($calculator->getArguments() as $name => $argument) {
            $builder->add($name, $this->getFormType($argument), $this->getFormTypeOptions($argument));
        }

        $builder
            ->add('preview', SubmitType::class, [
                'attr' => [
                    'formtarget' => 'preview',
                ],
            ])
            ->add('submit', SubmitType::class);

        return $builder->getForm();
    }

    private function getFormTypeOptions($argument)
    {
        $type = $argument['type'];
        $options = [
            'label' => $argument['name'] ?? $name,
            'required' => $argument['required'],
            'help' => $argument['description'] ?? null,
            'data' => $this->getData($argument),
        ];

        switch ($type) {
            case 'date':
                $options['widget'] = 'single_text';
                break;
        }

        return $options;
    }

    private function getFormType(array $argument): string
    {
        $type = $argument['type'];

        switch ($type) {
            case 'bool':
                return CheckboxType::class;
            case 'date':
                return DateType::class;
            case 'int':
                return IntegerType::class;
        }

        return TextType::class;
    }

    private function getData($argument)
    {
        $data = $argument['default'] ?? null;

        switch ($argument['type']) {
            case 'date':
                return new \DateTimeImmutable($data);
        }

        return $data;
    }
}
